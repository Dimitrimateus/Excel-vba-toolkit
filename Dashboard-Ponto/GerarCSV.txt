Attribute VB_Name = "GerarCSVPonto"
Option Explicit

' =====================================================================
' GerarCSVPonto
' =====================================================================
' O QUE ESTE MÓDULO FAZ, EM UMA FRASE:
' Lê várias abas de uma planilha de RH ("Tratamento Ponto") e escreve
' uma aba "CSV" (uma linha por ocorrência/registro) no formato que o
' Painel de Ponto (Dashboard-Ponto/index.html, um site estático que
' roda só no navegador) sabe importar e exibir em gráficos/tabelas.
'
' Este módulo é o único elo entre a planilha do RH e o painel web: o
' painel nunca lê Excel diretamente, só o CSV que sai daqui. Por isso,
' qualquer coluna nova que o painel precise mostrar tem que primeiro
' ganhar uma coluna correspondente aqui (ver PrepararAbaCSV) e ser
' preenchida em algum lugar deste código.
'
' -----------------------------------------------------------------
' AS DUAS MACROS (na ordem em que você roda):
'   1. GerarAbaCSV        - lê tudo, escreve a aba "CSV" dentro desta
'                            mesma planilha. Pode rodar quantas vezes
'                            quiser (sempre limpa e reescreve a aba).
'   2. ExportarCSVPorGestor - lê a aba "CSV" já pronta e salva arquivos
'                            .csv de verdade em disco: um por gestor
'                            (pra mandar só a parte de cada um) mais um
'                            "dados_TODOS.csv" com tudo (esse último é
'                            o que o RH sobe no painel).
'
' -----------------------------------------------------------------
' DE ONDE CADA COISA VEM (abas de entrada, todas na MESMA planilha):
'
'   "Tratamento" (obrigatória) — uma linha por ocorrência já revisada
'     pelo RH no mês atual (atraso, falta, hora extra, etc.). É a
'     fonte principal: define QUAIS dias viram ocorrência de hora
'     extra/falta (colunas "Extras"/"Faltas"/"Extra 100%"), a
'     situação, o status e o gestor/cargo do colaborador.
'
'   "RE 08.09" (obrigatória) — cadastro de colaboradores (matrícula →
'     setor "Nome Unidade" e cargo). Usada como lookup de apoio sempre
'     que a informação não está direto na aba de origem da linha.
'
'   "PontoNet" (opcional) — relatório de "Ausência de marcação" colado
'     nesta planilha (ou escolhido como arquivo separado, se essa aba
'     não existir). Não gera linhas novas no CSV: só enriquece a
'     situação da FALTA já detectada pela Tratamento e alimenta a
'     coluna "data_tratativa_pontonet" (usada pelo painel para calcular
'     a demora de tratativa). Se não for encontrada, a demora fica em
'     branco e o resto do CSV continua saindo normalmente.
'
'   "Cartão ponto ..." (uma ou mais; ex.: "Cartão ponto Julho",
'     "Cartão ponto Agosto", "Cartão ponto Atual" — qualquer texto
'     depois de "Cartão ponto" serve, o macro acha a aba pelo prefixo
'     do nome, não pelo nome exato) — export bruto do sistema de ponto,
'     um pouco diferente das outras abas: o cabeçalho de verdade fica
'     na LINHA 2 (a linha 1 só tem uns rótulos de grupo soltos, tipo
'     "Horas extras"), então toda leitura dessas abas passa
'     linhaHeader:=2 pra ColunaPorCabecalho. O macro usa até as 3 abas
'     desse tipo com a data mais recente (não precisa ficar renomeando
'     nada todo mês, só não deixar mais de 3 abas acumuladas). A mais
'     recente é tratada como "mês atual" e alimenta:
'       - BH/50%/100% do dia, pra completar horas de extra/falta que a
'         Tratamento já decidiu que existem (ver "Hora extra e falta:
'         de onde vem a QUANTIDADE" abaixo);
'       - "Hora Extra 100%" (linha própria, sempre que houver);
'       - "Ocorrências curtas" (linha própria, coluna "Tolerância <
'         15min", ver GerarOcorrenciasCurtasDoCartao);
'     Todas as abas encontradas (até 3) alimentam a auditoria de
'     "extra e falta no mesmo dia" dos últimos meses (ver
'     "Auditoria de 3 meses" abaixo).
'
'   "Pausas térmicas" (opcional) — aba já formatada pelo macro irmão
'     FormatarPausasTermicas.bas (pasta Pausas-Termicas/ do repo). Se
'     existir, sua seção "Totais por Colaborador" é repassada pro CSV
'     tal como está, sem recalcular nada (ver GerarResumoPausasTermicas).
'
' -----------------------------------------------------------------
' PRA ONDE CADA COISA VAI: A ABA "CSV" (formato de saída)
'
' Uma linha = um evento. O cabeçalho tem 22 colunas (ver
' PrepararAbaCSV); a maioria das linhas só preenche as primeiras 13
' (ocorrência normal — extra, falta, hora extra 100%, curta etc.); as
' 9 últimas só existem em 3 tipos de linha "resumo por colaborador"
' que não representam um dia específico (ver EscreverLinhaCSV):
'   - tipo_ocorrencia = "Auditoria Extra e Falta (3 Meses)" → preenche
'     ocorrencias_extra_falta_mes_atual / _3_meses.
'   - tipo_ocorrencia = "Resumo Pausas Térmicas" → preenche as 7
'     colunas pausas_.../trabalho_....
'   - qualquer outro tipo (Hora Extra, Falta, Atraso, Falta < 15min,
'     Extra < 15min, ...) → só usa as 13 primeiras colunas.
' O painel (index.html) sabe diferenciar esses 3 casos pelo próprio
' texto de tipo_ocorrencia (ver TIPOS_RESUMO/TIPOS_CURTAS/
' TIPOS_HORA_EXTRA lá no JS) e trata cada um numa vista própria, sem
' misturar com os gráficos/KPIs de ocorrência "normal".
'
' -----------------------------------------------------------------
' HORA EXTRA E FALTA: DE ONDE VEM A QUANTIDADE (o ponto mais sutil)
'
' As colunas "Extras"/"Faltas"/"Extra 100%" da Tratamento são só um
' GATE: decidem SE aquele dia vira uma linha de "Hora Extra" ou
' "Falta/Atraso" no CSV. A QUANTIDADE de minutos gravada, porém, vem
' preferencialmente do Cartão Ponto do mês atual (coluna "BH" pra
' falta coberta por banco de horas, "50%" pra hora extra normal) —
' e só cai de volta pro valor da própria Tratamento se não achar a
' linha correspondente no Cartão Ponto (pra nunca perder o dado). A
' hora extra a 100% (coluna "100%" do Cartão Ponto) é assunto à parte:
' sempre reportada em linha separada, com Check "S" ou "N", nunca
' somada à hora extra "normal". As colunas "60%"/"120%" não são usadas
' em lugar nenhum (pedido do RH).
'
' -----------------------------------------------------------------
' REGRAS DE EXCLUSÃO (quando uma linha da Tratamento NÃO vira nada)
'   - Situação = "" ou "Sem alteração" → não é ocorrência real, pulada
'     (troque INCLUIR_SEM_ALTERACAO pra True se isso mudar).
'   - Check = "S" → o sistema aponta uma ocorrência mas o RH já
'     confirmou que não é um problema real (ex.: 1 min de hora extra);
'     só a hora extra 100% dessa linha (se houver) ainda é reportada,
'     porque isso é sempre relevante independente do Check.
'
' -----------------------------------------------------------------
' AUDITORIA DE 3 MESES E OCORRÊNCIAS CURTAS (linhas geradas DEPOIS do
' loop principal, não linha a linha da Tratamento — ver o corpo de
' GerarAbaCSV logo após "Next r"):
'   - Extra+falta no mesmo dia: soma, por colaborador, quantos dias
'     tiveram "Sim" na coluna "Banco de horas e extra no mesmo dia" em
'     cada aba de Cartão Ponto encontrada (até 3 meses). O painel
'     destaca quem chega a 10+ no total dos 3 meses.
'   - Ocorrências curtas: agora vêm direto da coluna "Tolerância <
'     15min" do Cartão Ponto do mês atual (GerarOcorrenciasCurtasDoCartao),
'     cobrindo TODA linha com esse sinal — não depende mais do Check
'     da Tratamento (que só cobria um subconjunto menor).
'
' -----------------------------------------------------------------
' PONTOS QUE AINDA DEPENDEM DE UMA REGRA DE NEGÓCIO NÃO CONFIRMADA
' (procure "REGRA:" no código pra achar onde ajustar):
'   - Destino da hora extra (Banco de Horas x Pagamento): a Tratamento
'     não separa isso, então toda hora extra sai com destino em branco
'     (o painel trata "em branco" como Pagamento).
'   - Mapeamento Tratado/Tratando/Tratar → Pendente/Aprovado/
'     Reprovado/Regularizado (vocabulário do painel) é uma
'     simplificação: hoje só existe Regularizado (=Tratado) e Pendente
'     (tudo o mais). Ajuste MapearStatus se o RH quiser diferenciar
'     Aprovado/Reprovado de verdade.
'   - A coluna "BH" do Cartão Ponto aparece tanto em dias de "Falta
'     (Banco Horas)" (valor alto, o dia inteiro) quanto em dias
'     "Trabalhando" (valor baixo, minutos). Este código trata qualquer
'     valor de "BH" como falta coberta pelo banco. Se o "BH" de um dia
'     "Trabalhando" for crédito (e não falta), avise pra ajustar.
'
' -----------------------------------------------------------------
' Este módulo nunca foi testado dentro do Excel de verdade (foi
' escrito num ambiente sem Excel instalado, só validado simulando a
' mesma lógica em Python contra arquivos reais). Rode sempre numa
' CÓPIA da planilha antes de usar com dados de produção.
' =====================================================================

' ---- Regras configuráveis -------------------------------------------
Private Const INCLUIR_SEM_ALTERACAO As Boolean = False
' "S" = ocorrência sem problema real (confirmado pelo RH); mude para
' um array maior se algum dia existir um terceiro valor válido.
Private Const VALOR_CHECK_IGNORAR As String = "S"
' REGRA: sem informação separada de banco de horas x pagamento na aba
' Tratamento. Por padrão toda hora extra fica com destino em branco
' (o painel trata "em branco" como Pagamento). Ajuste aqui se preciso.
Private Const DESTINO_HORA_EXTRA_PADRAO As String = ""

' ---------------------------------------------------------------------
' Sub principal: gera a aba "CSV"
' ---------------------------------------------------------------------
' Roteiro desta Sub, em 4 fases (cada uma comentada no próprio código,
' procure pelos separadores "====="):
'   Fase 0 (abaixo): acha as abas de entrada e monta os dicionários de
'     apoio (ver "por que dicionários" logo mais abaixo).
'   Fase 1 (loop "For r = 2 To ultimaLinha"): uma linha da Tratamento
'     de cada vez, decide o que virar ocorrência e escreve no CSV.
'   Fase 2, 3, 4 (depois do "Next r"): três blocos independentes que
'     acrescentam linhas "resumo" ao MESMO CSV (auditoria de 3 meses,
'     ocorrências curtas, resumo de pausas térmicas) — nenhum dos três
'     depende do loop da Fase 1, só reaproveitam os dicionários já
'     montados (dictRE, dictGestorPorMatricula).
Public Sub GerarAbaCSV()
    Dim wbTrat As Workbook
    Set wbTrat = ThisWorkbook

    Dim wsTrat As Worksheet, wsRE As Worksheet, wsCartao As Worksheet
    Set wsTrat = wbTrat.Sheets("Tratamento")
    Set wsRE = wbTrat.Sheets("RE 08.09")

    ' Pode haver várias abas "Cartão ponto ..." (uma por mês, ex.: Julho/
    ' Agosto/Atual) — usamos até as 3 mais recentes (pela maior data na
    ' coluna DT, não pelo nome) para a auditoria de "extra e falta no
    ' mesmo dia" dos últimos meses. A mais recente delas também é a
    ' usada, como antes, para os valores de BH/50%/100%/Tolerância do
    ' mês atual.
    Const MAXIMO_MESES_AUDITORIA As Long = 3
    Dim abasCartao As Collection
    Set abasCartao = AbasMaisRecentes(EncontrarTodasAbas(wbTrat, "Cartão ponto"), MAXIMO_MESES_AUDITORIA)

    If abasCartao.Count > 0 Then
        Set wsCartao = abasCartao(1) ' já é a mais recente (AbasMaisRecentes ordena assim)
    Else
        Set wsCartao = Nothing
    End If

    ' Primeiro procura uma aba "PontoNet" na própria planilha; só abre
    ' um arquivo separado se essa aba não existir.
    Dim wbAus As Workbook
    Dim wsAus As Worksheet
    On Error Resume Next
    Set wsAus = wbTrat.Sheets("PontoNet")
    On Error GoTo 0

    If wsAus Is Nothing Then
        Set wbAus = AbrirArquivoAusencia()
        If Not wbAus Is Nothing Then Set wsAus = wbAus.Sheets(1)
    End If

    ' Por que dicionários (Scripting.Dictionary) em vez de procurar
    ' célula a célula toda vez? O loop principal, mais na frente, faz
    ' até 3 "buscas" por linha (RE, PontoNet, Cartão Ponto) — se cada
    ' busca varresse a aba inteira de novo, o tempo total cresceria
    ' multiplicando linhas-da-Tratamento × linhas-da-outra-aba (lento
    ' pra milhares de linhas). Em vez disso, cada aba é lida UMA vez
    ' aqui, isso monta um dicionário chave→dado (chave = matrícula, ou
    ' matrícula&"|"&data quando é por dia), e depois o loop principal só
    ' consulta esse dicionário (Dictionary.Exists/Item são O(1), não
    ' precisam varrer nada). colIdx é o mesmo princípio, mas pra nomes
    ' de coluna da própria Tratamento: dict("Nome do Cabeçalho") -> nº
    ' da coluna, montado uma vez, usado o loop inteiro.
    Dim dictRE As Object, dictAus As Object, dictCartao As Object, colIdx As Object
    Set dictRE = MontarDicionarioRE(wsRE)           ' matrícula -> (setor, cargo)
    Set dictAus = MontarDicionarioAusencia(wsAus)   ' matrícula|data -> (justificativa, avaliado em, integrado?)
    Set dictCartao = MontarDicionarioCartao(wsCartao) ' matrícula|data -> (BH, 50%, 100%) em minutos
    Set colIdx = MapearColunas(wsTrat)              ' "Nome do Cabeçalho" -> nº da coluna, na Tratamento

    Dim wsCSV As Worksheet
    Set wsCSV = PrepararAbaCSV(wbTrat)

    Dim ultimaLinha As Long
    ultimaLinha = wsTrat.Cells(wsTrat.Rows.Count, colIdx("Matricula")).End(xlUp).Row

    Dim linhaSaida As Long
    linhaSaida = 2

    ' guarda o primeiro gestor não vazio visto por matrícula, pra
    ' reaproveitar na linha de auditoria de 3 meses (colaboradores que só
    ' aparecem numa aba de Cartão Ponto antiga não têm gestor em nenhum
    ' outro lugar da planilha)
    Dim dictGestorPorMatricula As Object
    Set dictGestorPorMatricula = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To ultimaLinha

        ' Lê os campos "crus" desta linha da Tratamento. colIdx.Exists(...)
        ' é usado nas colunas opcionais (Check, Horas Excedentes) porque
        ' uma planilha mais antiga pode não ter essas colunas ainda — sem
        ' o Exists, colIdx("Check") lançaria erro de chave inexistente.
        Dim matricula As String, nome As String, gestor As String, situacao As String, statusTxt As String, checkTxt As String
        Dim horasExcedentesTxt As String
        Dim dataOcorrencia As Variant

        matricula = Trim$(CStr(wsTrat.Cells(r, colIdx("Matricula")).Value))
        nome = Trim$(CStr(wsTrat.Cells(r, colIdx("Nome")).Value))
        gestor = Trim$(CStr(wsTrat.Cells(r, colIdx("Gestor")).Value))
        situacao = Trim$(CStr(wsTrat.Cells(r, colIdx("Situação")).Value))
        statusTxt = Trim$(CStr(wsTrat.Cells(r, colIdx("Status")).Value))
        dataOcorrencia = wsTrat.Cells(r, colIdx("Data")).Value
        checkTxt = ""
        If colIdx.Exists("Check") Then checkTxt = Trim$(CStr(wsTrat.Cells(r, colIdx("Check")).Value))
        horasExcedentesTxt = ""
        If colIdx.Exists("Horas Excedentes") Then
            If Trim$(CStr(wsTrat.Cells(r, colIdx("Horas Excedentes")).Value)) <> "" Then horasExcedentesTxt = "Sim"
        End If

        ' Linha em branco (fim de fato dos dados, mesmo com ultimaLinha
        ' apontando mais longe) ou sem data válida: não dá pra fazer
        ' nada com ela, pula pra próxima.
        If nome = "" Or Not IsDate(dataOcorrencia) Then GoTo ProximaLinha

        If gestor <> "" And matricula <> "" Then
            If Not dictGestorPorMatricula.Exists(matricula) Then dictGestorPorMatricula.Add matricula, gestor
        End If

        ' Cartão Ponto: consultado sempre, mesmo em linhas Check = "S",
        ' porque a hora extra a 100% é reportada independente do Check
        ' (ocorrências curtas <15min vêm de outro lugar — ver
        ' GerarOcorrenciasCurtasDoCartao, chamada após este loop).
        Dim minFaltaBHCartao As Double, minExtra50Cartao As Double, minExtra100Cartao As Double
        ObterValoresCartao dictCartao, matricula, dataOcorrencia, minFaltaBHCartao, minExtra50Cartao, minExtra100Cartao

        ' Setor: só existe na "RE 08.09" (lookup por matrícula). Cargo
        ' tem duas fontes possíveis, nessa ordem de prioridade: se a
        ' própria Tratamento já tiver uma coluna "Cargo" preenchida
        ' (adicionada manualmente pelo RH), ela vence; senão cai pro
        ' cargo cadastrado na "RE 08.09".
        Dim setorNome As String, cargoNome As String, cargoTratamento As String
        ObterSetorCargo dictRE, matricula, setorNome, cargoNome
        cargoTratamento = ""
        If colIdx.Exists("Cargo") Then cargoTratamento = Trim$(CStr(wsTrat.Cells(r, colIdx("Cargo")).Value))
        If cargoTratamento <> "" Then cargoNome = cargoTratamento

        Dim situacaoTexto As String, avaliadoEm As Variant, temTratativa As Boolean
        ObterDadosAusencia dictAus, matricula, dataOcorrencia, situacaoTexto, avaliadoEm, temTratativa

        Dim statusFinal As String
        statusFinal = MapearStatus(statusTxt)

        Dim situacaoFinal As String
        situacaoFinal = situacao
        If situacaoTexto <> "" Then situacaoFinal = situacaoTexto

        ' Hora extra a 100%: reportada sempre que existir, em separado
        ' da hora extra "normal" (50%), Check = "S" ou não.
        If minExtra100Cartao > 0 Then
            EscreverLinhaCSV wsCSV, linhaSaida, dataOcorrencia, nome, matricula, gestor, setorNome, cargoNome, _
                "Hora Extra 100%", situacaoFinal, statusFinal, Round(minExtra100Cartao, 0), _
                DESTINO_HORA_EXTRA_PADRAO, Empty, horasExcedentesTxt
            linhaSaida = linhaSaida + 1
        End If

        If UCase$(checkTxt) = VALOR_CHECK_IGNORAR Then
            ' Check = "S": não é uma ocorrência real pro RH (extra/falta
            ' normal). "Ocorrências curtas" (<15min) não vêm mais daqui —
            ' agora são geradas direto da aba de Cartão Ponto do mês atual
            ' (ver GerarOcorrenciasCurtasDoCartao, chamada após este loop),
            ' cobrindo todas as linhas com "Tolerância < 15min" preenchida,
            ' não só as que também têm Check = "S" na Tratamento.
            GoTo ProximaLinha
        End If

        If Not INCLUIR_SEM_ALTERACAO And (situacao = "" Or situacao = "Sem alteração") Then GoTo ProximaLinha

        ' As colunas Extras/Faltas da aba Tratamento só decidem SE o dia
        ' entra como ocorrência de extra/falta (pedido do RH). A
        ' quantidade de horas gravada vem do Cartão Ponto (ver abaixo).
        Dim minExtraTrat As Double, minExtra100Trat As Double, minFaltaTrat As Double
        minExtraTrat = NzNum(wsTrat.Cells(r, colIdx("Extras")).Value) * 24 * 60
        minExtra100Trat = NzNum(wsTrat.Cells(r, colIdx("Extra 100%")).Value) * 24 * 60
        minFaltaTrat = NzNum(wsTrat.Cells(r, colIdx("Faltas")).Value) * 24 * 60
        Dim temExtraTrat As Boolean, temFaltaTrat As Boolean
        temExtraTrat = (minExtraTrat + minExtra100Trat) > 0
        temFaltaTrat = minFaltaTrat > 0

        ' Quantidade de horas "oficial" da hora extra "normal" (50%):
        ' Cartão Ponto, coluna S. Cai de volta pro valor da Tratamento
        ' se não achar a linha correspondente no Cartão Ponto. A hora
        ' extra a 100% (coluna U) já foi reportada em separado acima,
        ' então NÃO entra aqui, pra não contar em dobro.
        Dim minExtra As Double, minFalta As Double
        minExtra = IIf(minExtra50Cartao > 0, minExtra50Cartao, minExtraTrat + minExtra100Trat)
        minFalta = IIf(minFaltaBHCartao > 0, minFaltaBHCartao, minFaltaTrat)

        ' Linha de HORA EXTRA, se a aba Tratamento apontar extra nesse dia
        ' (gate). Não recebe a demora do PontoNet: essa demora é sobre a
        ' justificativa de ausência de marcação, não sobre a extra.
        If temExtraTrat Then
            EscreverLinhaCSV wsCSV, linhaSaida, dataOcorrencia, nome, matricula, gestor, setorNome, cargoNome, _
                "Hora Extra", situacaoFinal, statusFinal, Round(minExtra, 0), _
                DESTINO_HORA_EXTRA_PADRAO, Empty, horasExcedentesTxt
            linhaSaida = linhaSaida + 1
        End If

        ' Linha de FALTA/ATRASO, se a aba Tratamento apontar falta nesse
        ' dia (gate). Valor em minutos: Cartão Ponto (BH) ou Tratamento.
        If temFaltaTrat Then
            EscreverLinhaCSV wsCSV, linhaSaida, dataOcorrencia, nome, matricula, gestor, setorNome, cargoNome, _
                situacao, situacaoFinal, statusFinal, -Round(minFalta, 0), "", _
                IIf(temTratativa, avaliadoEm, Empty), horasExcedentesTxt
            linhaSaida = linhaSaida + 1
        End If

        ' Ocorrência sem impacto de horas (ex.: Registro duplo,
        ' Problema horário, Sem marcação já corrigida) - ainda entra
        ' no CSV com duracao_minutos = 0, só para contar nos gráficos
        ' de tipo/gestor/cargo/setor/data.
        If Not temExtraTrat And Not temFaltaTrat Then
            EscreverLinhaCSV wsCSV, linhaSaida, dataOcorrencia, nome, matricula, gestor, setorNome, cargoNome, _
                situacao, situacaoFinal, statusFinal, 0, "", _
                IIf(temTratativa, avaliadoEm, Empty), horasExcedentesTxt
            linhaSaida = linhaSaida + 1
        End If

ProximaLinha:
    Next r

    ' =====================================================================
    ' Auditoria de "extra e falta no mesmo dia" nos últimos meses
    ' ---------------------------------------------------------------------
    ' Uma linha por colaborador (além das linhas de ocorrência normais
    ' acima), somando quantos dias com "Banco de horas e extra no mesmo
    ' dia" = Sim cada um teve em cada aba de Cartão Ponto encontrada
    ' (até MAXIMO_MESES_AUDITORIA, a mais recente sendo o "mês atual").
    ' A partir de 10 no total de 3 meses, o painel sinaliza a linha para
    ' a auditoria (pedido do RH) — ver colunas
    ' ocorrencias_extra_falta_mes_atual / ocorrencias_extra_falta_3_meses.
    ' =====================================================================
    Dim qtdColaboradoresAuditoria As Long
    qtdColaboradoresAuditoria = 0

    If abasCartao.Count > 0 Then
        Dim dictNomesCartao As Object
        Set dictNomesCartao = CreateObject("Scripting.Dictionary")

        ' matricula -> Dictionary(nome da aba -> contagem de "Sim")
        Dim dictContagens As Object
        Set dictContagens = CreateObject("Scripting.Dictionary")

        Dim nomeAbaAtual As String
        nomeAbaAtual = wsCartao.Name ' abasCartao(1), a mais recente

        ' calculado uma única vez aqui fora: é sempre o mesmo valor pra
        ' toda linha de auditoria, não precisa (nem deve) ser recalculado
        ' a cada colaborador do loop "For Each matAud" abaixo — cada
        ' chamada a DataMaximaAba varre a coluna DT inteira da aba.
        Dim dataAtualCartao As Date
        dataAtualCartao = DataMaximaAba(wsCartao)

        ' Passo 1: pra cada aba de Cartão Ponto (até 3), conta quantos
        ' "Sim" cada matrícula teve NAQUELA aba, e guarda isso dentro de
        ' dictContagens como matrícula -> (nome da aba -> contagem).
        ' Esse "dicionário dentro de dicionário" é o que permite, no
        ' passo 2, separar "quanto foi só no mês atual" de "quanto foi
        ' no total dos 3 meses" pra cada colaborador.
        Dim abaIter As Variant
        For Each abaIter In abasCartao
            Dim wsIter As Worksheet
            Set wsIter = abaIter

            Dim dictSim As Object
            Set dictSim = ContarExtraFaltaMesmoDia(wsIter, dictNomesCartao)

            Dim matIter As Variant
            For Each matIter In dictSim.Keys
                If Not dictContagens.Exists(CStr(matIter)) Then
                    dictContagens.Add CStr(matIter), CreateObject("Scripting.Dictionary")
                End If
                dictContagens(CStr(matIter)).Add wsIter.Name, dictSim(matIter)
            Next matIter
        Next abaIter

        ' Passo 2: agora com os dados dos 3 meses todos já reunidos por
        ' matrícula, soma tudo (total3Meses) e separa à parte a
        ' contagem da aba mais recente (mesAtualCount) pra virar as duas
        ' colunas que o CSV precisa.
        Dim matAud As Variant
        For Each matAud In dictContagens.Keys
            Dim contagensAba As Object
            Set contagensAba = dictContagens(CStr(matAud))

            Dim total3Meses As Long, mesAtualCount As Long, nomeAbaK As Variant
            total3Meses = 0
            mesAtualCount = 0
            For Each nomeAbaK In contagensAba.Keys
                total3Meses = total3Meses + contagensAba(nomeAbaK)
                If CStr(nomeAbaK) = nomeAbaAtual Then mesAtualCount = contagensAba(nomeAbaK)
            Next nomeAbaK

            If total3Meses > 0 Then
                Dim setorAud As String, cargoAud As String, nomeAud As String, gestorAud As String
                ObterSetorCargo dictRE, CStr(matAud), setorAud, cargoAud

                nomeAud = ""
                If dictNomesCartao.Exists(CStr(matAud)) Then nomeAud = dictNomesCartao(CStr(matAud))
                If nomeAud = "" Then nomeAud = "Matrícula " & matAud

                gestorAud = ""
                If dictGestorPorMatricula.Exists(CStr(matAud)) Then gestorAud = dictGestorPorMatricula(CStr(matAud))

                EscreverLinhaCSV wsCSV, linhaSaida, dataAtualCartao, nomeAud, CStr(matAud), gestorAud, _
                    setorAud, cargoAud, "Auditoria Extra e Falta (3 Meses)", "", "Pendente", 0, "", Empty, "", _
                    mesAtualCount, total3Meses
                linhaSaida = linhaSaida + 1
                qtdColaboradoresAuditoria = qtdColaboradoresAuditoria + 1
            End If
        Next matAud
    End If

    ' =====================================================================
    ' Ocorrências curtas (< 15 minutos), direto da aba de Cartão Ponto do
    ' mês atual — pedido do RH: usar a coluna "Tolerância < 15min" do
    ' Cartão Ponto Atual diretamente, cobrindo TODAS as linhas com esse
    ' sinal (não só as que também batem Check="S" na Tratamento, que é
    ' um subconjunto menor).
    ' =====================================================================
    Dim qtdCurtas As Long
    qtdCurtas = 0
    If Not wsCartao Is Nothing Then
        GerarOcorrenciasCurtasDoCartao wsCartao, dictRE, dictGestorPorMatricula, wsCSV, linhaSaida, qtdCurtas
    End If

    ' =====================================================================
    ' Resumo de Pausas Térmicas por colaborador (se a aba existir)
    ' ---------------------------------------------------------------------
    ' Usa a seção "Totais por Colaborador" que o macro FormatarPausasTermicas
    ' já deixa pronta na aba "Pausas térmicas" (uma linha por colaborador,
    ' com Pausa corretas/menor/maior e Trabalho correto/maior/menor 1:40 e
    ' Marcações Ímpares) — não recalcula nada, só repassa esses números
    ' pro CSV, numa linha própria por colaborador.
    ' =====================================================================
    Dim qtdResumoPausas As Long
    qtdResumoPausas = 0
    Dim wsPausas As Worksheet
    On Error Resume Next
    Set wsPausas = wbTrat.Sheets("Pausas térmicas")
    On Error GoTo 0
    If Not wsPausas Is Nothing Then
        ' dataAtualCartao só foi calculada acima se wsCartao existe (é o
        ' mesmo teste); sem Cartão Ponto nenhum, usa a data de hoje como
        ' referência só pra a linha ter uma data válida no CSV.
        GerarResumoPausasTermicas wsPausas, dictRE, dictGestorPorMatricula, wsCSV, linhaSaida, _
            IIf(Not wsCartao Is Nothing, dataAtualCartao, Date), qtdResumoPausas
    End If

    If Not wbAus Is Nothing Then wbAus.Close SaveChanges:=False

    Dim msgAuditoria As String
    If abasCartao.Count > 0 Then
        msgAuditoria = vbCrLf & qtdColaboradoresAuditoria & " colaborador(es) com extra+falta no mesmo dia nos últimos " & _
            abasCartao.Count & " mês(es) (" & nomeAbaAtual & " = mês atual)."
    Else
        msgAuditoria = vbCrLf & "Nenhuma aba ""Cartão ponto ..."" encontrada; auditoria de 3 meses não gerada."
    End If
    msgAuditoria = msgAuditoria & vbCrLf & qtdCurtas & " ocorrência(s) curta(s) (<15min) direto do Cartão Ponto."
    msgAuditoria = msgAuditoria & vbCrLf & IIf(wsPausas Is Nothing, _
        "Aba ""Pausas térmicas"" não encontrada; resumo de pausas térmicas não gerado.", _
        qtdResumoPausas & " colaborador(es) no resumo de Pausas Térmicas.")

    MsgBox (linhaSaida - 2) & " linha(s) geradas na aba CSV" & msgAuditoria, vbInformation, "Gerar CSV"
End Sub

' ---------------------------------------------------------------------
' Sub: exporta a aba "CSV" em um arquivo por gestor + um arquivo geral
' ---------------------------------------------------------------------
Public Sub ExportarCSVPorGestor()
    Dim wsCSV As Worksheet
    On Error Resume Next
    Set wsCSV = ThisWorkbook.Sheets("CSV")
    On Error GoTo 0
    If wsCSV Is Nothing Then
        MsgBox "Não encontrei a aba 'CSV'. Rode primeiro a macro GerarAbaCSV.", vbExclamation
        Exit Sub
    End If

    Dim ultimaLinha As Long
    ultimaLinha = wsCSV.Cells(wsCSV.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "A aba 'CSV' está vazia. Rode primeiro a macro GerarAbaCSV.", vbExclamation
        Exit Sub
    End If

    Dim pasta As String
    pasta = EscolherPasta()
    If pasta = "" Then Exit Sub

    Const COL_GESTOR As Long = 4 ' ordem fixa: data,colaborador,matricula,gestor,...

    Dim dictGestores As Object
    Set dictGestores = CreateObject("Scripting.Dictionary")

    Dim r As Long, g As String
    For r = 2 To ultimaLinha
        g = Trim$(CStr(wsCSV.Cells(r, COL_GESTOR).Value))
        If g <> "" And Not dictGestores.Exists(g) Then dictGestores.Add g, 1
    Next r

    Dim k As Variant
    For Each k In dictGestores.Keys
        ExportarLinhasParaCSV wsCSV, pasta & "\dados_" & NomeArquivoSeguro(CStr(k)) & ".csv", CStr(k)
    Next k

    ExportarLinhasParaCSV wsCSV, pasta & "\dados_TODOS.csv", ""

    MsgBox dictGestores.Count & " arquivo(s) por gestor + 1 arquivo geral (dados_TODOS.csv) gerados em:" _
        & vbNewLine & pasta, vbInformation, "Exportar CSV"
End Sub

' =====================================================================
' Funções auxiliares
' =====================================================================

' Pede ao usuário o arquivo PontoNet de "ausência de marcação" (workbook
' separado do principal, com a "demora" — quanto tempo entre a
' ocorrência e a tratativa do gestor). É opcional: se o usuário cancelar
' o diálogo, devolve Nothing e o chamador segue em frente sem esse dado
' (ObterDadosAusencia trata dictAus vazio normalmente).
Private Function AbrirArquivoAusencia() As Workbook
    Dim caminho As Variant
    caminho = Application.GetOpenFilename( _
        FileFilter:="Planilhas Excel (*.xlsx;*.xls),*.xlsx;*.xls", _
        Title:="Selecione o arquivo 'Ausência de marcação' do PontoNet (ou cancele para pular a demora)")
    If VarType(caminho) = vbBoolean Then
        Set AbrirArquivoAusencia = Nothing
    Else
        Set AbrirArquivoAusencia = Workbooks.Open(CStr(caminho), ReadOnly:=True)
    End If
End Function

' Acha TODAS as abas cujo nome comece com "prefixo" (ex.: "Cartão ponto
' Julho", "Cartão ponto Agosto", "Cartão ponto Atual"). Usado para a
' auditoria de "extra e falta no mesmo dia" nos últimos meses, que
' precisa somar todas as abas de Cartão Ponto presentes na planilha —
' não só a mais recente.
Private Function EncontrarTodasAbas(wb As Workbook, prefixo As String) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If LCase$(Left$(Trim$(ws.Name), Len(prefixo))) = LCase$(prefixo) Then
            col.Add ws
        End If
    Next ws
    Set EncontrarTodasAbas = col
End Function

' Maior data encontrada na coluna "DT" da aba (0 = aba vazia/sem
' coluna DT). Usado para descobrir qual aba de Cartão Ponto é a do mês
' atual (a de data mais recente), sem depender do nome da aba.
Private Function DataMaximaAba(ws As Worksheet) As Date
    Dim colData As Long, ultimaLinha As Long, r As Long, dt As Variant
    Dim maxData As Date
    maxData = 0
    If ws Is Nothing Then
        DataMaximaAba = maxData
        Exit Function
    End If
    colData = ColunaPorCabecalho(ws, "DT", 2)
    If colData = 0 Then
        DataMaximaAba = maxData
        Exit Function
    End If
    ultimaLinha = ws.Cells(ws.Rows.Count, colData).End(xlUp).Row
    For r = 3 To ultimaLinha
        dt = ws.Cells(r, colData).Value
        If IsDate(dt) Then
            If CDate(dt) > maxData Then maxData = CDate(dt)
        End If
    Next r
    DataMaximaAba = maxData
End Function

' Dentre as abas de Cartão Ponto encontradas, mantém só as N com a
' data mais recente (evita que uma aba antiga esquecida na planilha
' entre pra sempre no cálculo de 3 meses). Devolve uma NOVA Collection,
' ordenada da mais recente para a mais antiga.
Private Function AbasMaisRecentes(abas As Collection, maximoAbas As Long) As Collection
    Dim n As Long
    n = abas.Count
    Dim wsArr() As Worksheet, dtArr() As Date
    If n = 0 Then
        Set AbasMaisRecentes = New Collection
        Exit Function
    End If
    ReDim wsArr(1 To n)
    ReDim dtArr(1 To n)

    Dim i As Long, a As Variant
    i = 1
    For Each a In abas
        Set wsArr(i) = a
        dtArr(i) = DataMaximaAba(wsArr(i))
        i = i + 1
    Next a

    ' ordenação simples (poucas abas, não precisa de nada sofisticado)
    Dim j As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If dtArr(j) > dtArr(i) Then
                Dim tmpD As Date, tmpW As Worksheet
                tmpD = dtArr(i): dtArr(i) = dtArr(j): dtArr(j) = tmpD
                Set tmpW = wsArr(i): Set wsArr(i) = wsArr(j): Set wsArr(j) = tmpW
            End If
        Next j
    Next i

    Dim resultado As New Collection
    For i = 1 To n
        If i > maximoAbas Then Exit For
        resultado.Add wsArr(i)
    Next i
    Set AbasMaisRecentes = resultado
End Function

' Conta, por matrícula, quantas linhas têm "Sim" na coluna "Banco de
' horas e extra no mesmo dia" dessa aba de Cartão Ponto (auditoria de
' extra+falta no mesmo dia). Também preenche dictNomes (matricula ->
' nome) para colaboradores que só existem nessa aba antiga (não estão
' mais na Tratamento do mês atual).
Private Function ContarExtraFaltaMesmoDia(ws As Worksheet, dictNomes As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set ContarExtraFaltaMesmoDia = dict
        Exit Function
    End If

    Dim colMat As Long, colNome As Long, colFlag As Long
    colMat = ColunaPorCabecalho(ws, "Matrícula", 2)
    colNome = ColunaPorCabecalho(ws, "Nome", 2)
    colFlag = ColunaPorCabecalho(ws, "Banco de horas e extra no mesmo dia", 2)
    If colMat = 0 Or colFlag = 0 Then
        Set ContarExtraFaltaMesmoDia = dict
        Exit Function
    End If

    Dim ultimaLinha As Long, r As Long, mat As String, flagTxt As String, nome As String
    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 3 To ultimaLinha
        mat = Trim$(CStr(ws.Cells(r, colMat).Value))
        If mat <> "" Then
            flagTxt = LCase$(Trim$(CStr(ws.Cells(r, colFlag).Value)))
            If flagTxt = "sim" Then
                If dict.Exists(mat) Then
                    dict(mat) = dict(mat) + 1
                Else
                    dict(mat) = 1
                End If
            End If
            If Not dictNomes.Exists(mat) Then
                nome = IIf(colNome > 0, Trim$(CStr(ws.Cells(r, colNome).Value)), "")
                If nome <> "" Then dictNomes.Add mat, nome
            End If
        End If
    Next r
    Set ContarExtraFaltaMesmoDia = dict
End Function

' Gera, na aba CSV, uma linha de ocorrência para cada linha da aba de
' Cartão Ponto do mês atual que tiver a coluna "Tolerância < 15min"
' preenchida ("Falta < 15min", "Extra < 15min" ou "Extra e Falta <
' 15min") — direto da fonte, sem depender de Check="S" na Tratamento.
' Uma linha "Extra e Falta < 15min" gera DUAS linhas no CSV (uma de
' falta, uma de extra), porque são dois eventos distintos no mesmo dia.
Private Sub GerarOcorrenciasCurtasDoCartao(ws As Worksheet, dictRE As Object, dictGestorPorMatricula As Object, _
    wsCSV As Worksheet, ByRef linhaSaida As Long, ByRef qtdGerada As Long)

    Dim colMat As Long, colNome As Long, colData As Long, colBH As Long
    Dim col50 As Long, col100 As Long, colTolerancia As Long
    colMat = ColunaPorCabecalho(ws, "Matrícula", 2)
    colNome = ColunaPorCabecalho(ws, "Nome", 2)
    colData = ColunaPorCabecalho(ws, "DT", 2)
    colBH = ColunaPorCabecalho(ws, "BH", 2)
    col50 = ColunaPorCabecalho(ws, "50%", 2)
    col100 = ColunaPorCabecalho(ws, "100%", 2)
    colTolerancia = ColunaPorCabecalho(ws, "Tolerância < 15min", 2)
    If colMat = 0 Or colData = 0 Or colTolerancia = 0 Then Exit Sub

    Dim ultimaLinha As Long, r As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 3 To ultimaLinha
        Dim tol As String
        tol = Trim$(CStr(ws.Cells(r, colTolerancia).Value))
        If tol = "" Then GoTo ProximaLinhaCurta

        Dim mat As String, dt As Variant, nome As String
        mat = Trim$(CStr(ws.Cells(r, colMat).Value))
        dt = ws.Cells(r, colData).Value
        If mat = "" Or Not IsDate(dt) Then GoTo ProximaLinhaCurta

        nome = IIf(colNome > 0, Trim$(CStr(ws.Cells(r, colNome).Value)), "")
        If nome = "" Then nome = "Matrícula " & mat

        Dim setorNome As String, cargoNome As String
        ObterSetorCargo dictRE, mat, setorNome, cargoNome

        Dim gestor As String
        gestor = ""
        If dictGestorPorMatricula.Exists(mat) Then gestor = dictGestorPorMatricula(mat)

        Dim minFaltaBH As Double, minExtra As Double
        minFaltaBH = IIf(colBH > 0, NzNum(ws.Cells(r, colBH).Value), 0) * 24 * 60
        minExtra = (IIf(col50 > 0, NzNum(ws.Cells(r, col50).Value), 0) _
                  + IIf(col100 > 0, NzNum(ws.Cells(r, col100).Value), 0)) * 24 * 60

        If InStr(1, tol, "Falta", vbTextCompare) > 0 Then
            EscreverLinhaCSV wsCSV, linhaSaida, dt, nome, mat, gestor, setorNome, cargoNome, _
                "Falta < 15min", "Falta < 15min", "Pendente", -Round(minFaltaBH, 0), "", Empty
            linhaSaida = linhaSaida + 1
            qtdGerada = qtdGerada + 1
        End If
        If InStr(1, tol, "Extra", vbTextCompare) > 0 Then
            EscreverLinhaCSV wsCSV, linhaSaida, dt, nome, mat, gestor, setorNome, cargoNome, _
                "Extra < 15min", "Extra < 15min", "Pendente", Round(minExtra, 0), DESTINO_HORA_EXTRA_PADRAO, Empty
            linhaSaida = linhaSaida + 1
            qtdGerada = qtdGerada + 1
        End If

ProximaLinhaCurta:
    Next r
End Sub

' Acha a linha de cabeçalho "Matrícula / ... / Pausa corretas / ..." da
' seção "Totais por Colaborador" que o macro FormatarPausasTermicas
' deixa pronta na aba "Pausas térmicas" (0 se não achar).
Private Function LocalizarCabecalhoTotaisPausas(ws As Worksheet) As Long
    Dim r As Long, ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To ultimaLinha
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "Matrícula" And _
           Trim$(CStr(ws.Cells(r, 4).Value)) = "Pausa corretas" Then
            LocalizarCabecalhoTotaisPausas = r
            Exit Function
        End If
    Next r
    LocalizarCabecalhoTotaisPausas = 0
End Function

' Repassa pro CSV, uma linha por colaborador, os números já calculados
' pelo FormatarPausasTermicas na seção "Totais por Colaborador" da aba
' "Pausas térmicas" (não recalcula nada, só lê e copia).
Private Sub GerarResumoPausasTermicas(ws As Worksheet, dictRE As Object, dictGestorPorMatricula As Object, _
    wsCSV As Worksheet, ByRef linhaSaida As Long, dataRef As Variant, ByRef qtdGerada As Long)

    Dim linhaCabecalho As Long
    linhaCabecalho = LocalizarCabecalhoTotaisPausas(ws)
    If linhaCabecalho = 0 Then Exit Sub

    Dim r As Long
    r = linhaCabecalho + 1
    Do While Trim$(CStr(ws.Cells(r, 1).Value)) <> ""
        Dim mat As String, nome As String, cargoPausas As String
        mat = Trim$(CStr(ws.Cells(r, 1).Value))
        nome = Trim$(CStr(ws.Cells(r, 2).Value))
        cargoPausas = Trim$(CStr(ws.Cells(r, 3).Value))

        Dim setorNome As String, cargoRE As String
        ObterSetorCargo dictRE, mat, setorNome, cargoRE
        If cargoPausas = "" Then cargoPausas = cargoRE

        Dim gestor As String
        gestor = ""
        If dictGestorPorMatricula.Exists(mat) Then gestor = dictGestorPorMatricula(mat)

        EscreverLinhaCSV wsCSV, linhaSaida, dataRef, nome, mat, gestor, setorNome, cargoPausas, _
            "Resumo Pausas Térmicas", "", "Pendente", 0, "", Empty, _
            pausaCorretas:=CLng(NzNum(ws.Cells(r, 4).Value)), _
            pausaMenor20:=CLng(NzNum(ws.Cells(r, 5).Value)), _
            pausaMaior20:=CLng(NzNum(ws.Cells(r, 6).Value)), _
            trabalhoCorreto140:=CLng(NzNum(ws.Cells(r, 7).Value)), _
            trabalhoMaior140:=CLng(NzNum(ws.Cells(r, 8).Value)), _
            trabalhoMenor140:=CLng(NzNum(ws.Cells(r, 9).Value)), _
            pausasMarcacoesImpares:=CLng(NzNum(ws.Cells(r, 10).Value))
        linhaSaida = linhaSaida + 1
        qtdGerada = qtdGerada + 1

        r = r + 1
    Loop
End Sub

' linhaHeader: linha onde está o cabeçalho (1 por padrão). As abas de
' Cartão Ponto têm o cabeçalho de verdade na LINHA 2 (a linha 1 só tem
' uns rótulos de grupo esparsos, tipo "Horas extras"), então quem
' chamar essa função pra uma aba de Cartão Ponto precisa passar
' linhaHeader:=2.
Private Function ColunaPorCabecalho(ws As Worksheet, cabecalho As String, Optional linhaHeader As Long = 1) As Long
    Dim ultimaColuna As Long, c As Long
    ultimaColuna = ws.Cells(linhaHeader, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultimaColuna
        If Trim$(CStr(ws.Cells(linhaHeader, c).Value)) = cabecalho Then
            ColunaPorCabecalho = c
            Exit Function
        End If
    Next c
    ColunaPorCabecalho = 0
End Function

' Igual ColunaPorCabecalho, mas monta o mapa "nome da coluna -> número"
' inteiro de uma vez (cabeçalho sempre na linha 1 aqui — é só usada com
' a aba Tratamento, que segue o padrão normal de cabeçalho). Usada no
' lugar de várias chamadas a ColunaPorCabecalho porque a Tratamento tem
' muito mais colunas lidas por linha (12+) do que as outras abas.
Private Function MapearColunas(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim ultimaColuna As Long, c As Long, h As String
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultimaColuna
        h = Trim$(CStr(ws.Cells(1, c).Value))
        If h <> "" And Not dict.Exists(h) Then dict.Add h, c
    Next c
    Set MapearColunas = dict
End Function

' Cadastro de colaboradores (aba "RE 08.09"): matrícula -> (setor,
' cargo). "Nome Unidade" na RE 08.09 é o que o resto do código chama
' de "setor" (o painel usa esse termo). Consultado por ObterSetorCargo
' logo abaixo sempre que uma matrícula precisa de setor/cargo e a aba
' de origem não trouxe essa informação junto.
Private Function MontarDicionarioRE(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim colMat As Long, colUnidadeNome As Long, colCargo As Long
    colMat = ColunaPorCabecalho(ws, "Matrícula")
    colUnidadeNome = ColunaPorCabecalho(ws, "Nome Unidade")
    colCargo = ColunaPorCabecalho(ws, "Cargo")
    If colMat = 0 Then
        Set MontarDicionarioRE = dict
        Exit Function
    End If

    Dim ultimaLinha As Long, r As Long, mat As String
    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 2 To ultimaLinha
        mat = Trim$(CStr(ws.Cells(r, colMat).Value))
        If mat <> "" And Not dict.Exists(mat) Then
            dict.Add mat, Array( _
                IIf(colUnidadeNome > 0, Trim$(CStr(ws.Cells(r, colUnidadeNome).Value)), ""), _
                IIf(colCargo > 0, Trim$(CStr(ws.Cells(r, colCargo).Value)), ""))
        End If
    Next r
    Set MontarDicionarioRE = dict
End Function

' Lookup simples no dicionário montado por MontarDicionarioRE. ByRef
' porque o chamador já tem setorNome/cargoNome com valores default (ou
' vindos da Tratamento) e só quer sobrescrevê-los quando a matrícula
' existir na RE 08.09 — matrícula não encontrada não altera nada.
Private Sub ObterSetorCargo(dictRE As Object, matricula As String, ByRef setorNome As String, ByRef cargoNome As String)
    If dictRE.Exists(matricula) Then
        Dim arr As Variant
        arr = dictRE(matricula)
        setorNome = arr(0)
        cargoNome = arr(1)
        If setorNome = "" Then setorNome = "Sem setor"
    Else
        setorNome = "Sem setor"
        cargoNome = ""
    End If
End Sub

' Chave: Matricula & "|" & AAAA-MM-DD. Guarda, em minutos, a coluna
' "BH" (falta coberta por banco de horas, coluna P), a coluna "50%"
' (coluna S, hora extra "normal") e a coluna "100%" (coluna U, hora
' extra a 100%) — "60%" e "120%" não entram, por pedido do RH. Se a aba
' não existir (não foi encontrada na planilha), devolve um dicionário
' vazio e o chamador cai de volta para os valores da aba Tratamento.
' (A coluna "Tolerância < 15min" NÃO é lida aqui: "ocorrências curtas"
' tem sua própria função dedicada, GerarOcorrenciasCurtasDoCartao, que
' lê essa coluna direto — ver mais abaixo.)
Private Function MontarDicionarioCartao(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set MontarDicionarioCartao = dict
        Exit Function
    End If

    Dim colMat As Long, colData As Long, colBH As Long
    Dim col50 As Long, col100 As Long
    colMat = ColunaPorCabecalho(ws, "Matrícula", 2)
    colData = ColunaPorCabecalho(ws, "DT", 2)
    colBH = ColunaPorCabecalho(ws, "BH", 2)
    col50 = ColunaPorCabecalho(ws, "50%", 2)
    col100 = ColunaPorCabecalho(ws, "100%", 2)
    If colMat = 0 Or colData = 0 Then
        Set MontarDicionarioCartao = dict
        Exit Function
    End If

    Dim ultimaLinha As Long, r As Long
    Dim mat As String, dt As Variant, chave As String
    Dim bhMin As Double, extra50Min As Double, extra100Min As Double

    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 3 To ultimaLinha
        mat = Trim$(CStr(ws.Cells(r, colMat).Value))
        dt = ws.Cells(r, colData).Value
        If mat <> "" And IsDate(dt) Then
            chave = mat & "|" & Format(CDate(dt), "yyyy-mm-dd")
            bhMin = IIf(colBH > 0, NzNum(ws.Cells(r, colBH).Value), 0) * 24 * 60
            extra50Min = IIf(col50 > 0, NzNum(ws.Cells(r, col50).Value), 0) * 24 * 60
            extra100Min = IIf(col100 > 0, NzNum(ws.Cells(r, col100).Value), 0) * 24 * 60
            ' se a matrícula tiver mais de uma linha no mesmo dia, soma os minutos
            If dict.Exists(chave) Then
                Dim antigo As Variant
                antigo = dict(chave)
                dict(chave) = Array(antigo(0) + bhMin, antigo(1) + extra50Min, antigo(2) + extra100Min)
            Else
                dict(chave) = Array(bhMin, extra50Min, extra100Min)
            End If
        End If
    Next r
    Set MontarDicionarioCartao = dict
End Function

' Lookup no dicionário montado por MontarDicionarioCartao. Sempre zera
' os três ByRef antes de procurar, então "matrícula/data não encontrada
' no Cartão Ponto" e "encontrada mas com 0 minutos" dão o mesmo
' resultado pro chamador — o que é o comportamento certo aqui, já que
' o valor em minutos é tudo que os três casos têm em comum.
Private Sub ObterValoresCartao(dictCartao As Object, matricula As String, dataOcorrencia As Variant, _
    ByRef minFaltaBH As Double, ByRef minExtra50 As Double, ByRef minExtra100 As Double)

    minFaltaBH = 0
    minExtra50 = 0
    minExtra100 = 0
    If dictCartao Is Nothing Then Exit Sub
    If Not IsDate(dataOcorrencia) Then Exit Sub

    Dim chave As String
    chave = matricula & "|" & Format(CDate(dataOcorrencia), "yyyy-mm-dd")
    If dictCartao.Exists(chave) Then
        Dim arr As Variant
        arr = dictCartao(chave)
        minFaltaBH = arr(0)
        minExtra50 = arr(1)
        minExtra100 = arr(2)
    End If
End Sub

' Chave: Matricula & "|" & AAAA-MM-DD. Guarda a justificativa em
' texto, a data "Avaliado em:" e se o caso já está "INTEGRADO"
' (fechado). Só usamos "Avaliado em:" quando INTEGRADO = True, porque
' nas linhas ainda em aberto ("AGUARD. JUSTIFICATIVA" / "AGUARD.
' ANALISE GESTOR") esse campo vem com um valor provisório que não
' representa a demora real (confirmado comparando várias linhas do
' relatório de exemplo).
Private Function MontarDicionarioAusencia(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set MontarDicionarioAusencia = dict
        Exit Function
    End If

    Dim colMat As Long, colData As Long, colSituacao As Long, colJustificativa As Long, colAvaliado As Long
    colMat = ColunaPorCabecalho(ws, "Matrícula")
    colData = ColunaPorCabecalho(ws, "Data falta")
    colSituacao = ColunaPorCabecalho(ws, "Situação atual")
    colJustificativa = ColunaPorCabecalho(ws, "Justificativa")
    colAvaliado = ColunaPorCabecalho(ws, "Avaliado em:")
    If colMat = 0 Or colData = 0 Then
        Set MontarDicionarioAusencia = dict
        Exit Function
    End If

    Dim ultimaLinha As Long, r As Long
    Dim mat As String, dt As Variant, chave As String
    Dim situacaoAtual As String, integrado As Boolean

    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 2 To ultimaLinha
        mat = Trim$(CStr(ws.Cells(r, colMat).Value))
        dt = ws.Cells(r, colData).Value
        If mat <> "" And IsDate(dt) Then
            chave = mat & "|" & Format(CDate(dt), "yyyy-mm-dd")
            situacaoAtual = Trim$(CStr(ws.Cells(r, colSituacao).Value))
            integrado = (UCase$(situacaoAtual) = "INTEGRADO")
            dict(chave) = Array( _
                IIf(colJustificativa > 0, Trim$(CStr(ws.Cells(r, colJustificativa).Value)), ""), _
                IIf(colAvaliado > 0, ws.Cells(r, colAvaliado).Value, Empty), _
                integrado)
        End If
    Next r
    Set MontarDicionarioAusencia = dict
End Function

' Lookup no dicionário montado por MontarDicionarioAusencia (arquivo
' PontoNet separado e opcional — ver AbrirArquivoAusencia). temTratativa
' só vira True quando a linha tem status "integrado" E uma data válida
' de avaliação: é essa combinação que garante que avaliadoEm representa
' a demora real, e não um placeholder de linha ainda em aberto.
Private Sub ObterDadosAusencia(dictAus As Object, matricula As String, dataOcorrencia As Variant, _
    ByRef situacaoTexto As String, ByRef avaliadoEm As Variant, ByRef temTratativa As Boolean)

    situacaoTexto = ""
    avaliadoEm = Empty
    temTratativa = False
    If dictAus Is Nothing Then Exit Sub
    If Not IsDate(dataOcorrencia) Then Exit Sub

    Dim chave As String
    chave = matricula & "|" & Format(CDate(dataOcorrencia), "yyyy-mm-dd")
    If dictAus.Exists(chave) Then
        Dim arr As Variant
        arr = dictAus(chave)
        situacaoTexto = arr(0)
        If arr(2) = True And IsDate(arr(1)) Then
            avaliadoEm = arr(1)
            temTratativa = True
        End If
    End If
End Sub

' REGRA: ajuste este mapeamento se o vocabulário de status do painel
' precisar refletir Tratado/Tratando/Tratar tal como estão, em vez de
' Pendente/Aprovado/Reprovado/Regularizado.
Private Function MapearStatus(txt As String) As String
    Select Case UCase$(Trim$(txt))
        Case "TRATADO"
            MapearStatus = "Regularizado"
        Case "TRATANDO", "TRATAR"
            MapearStatus = "Pendente"
        Case Else
            MapearStatus = "Pendente"
    End Select
End Function

' "Null-zero": converte pra número com segurança. Células vazias ou
' com texto (ex.: "-") não são numéricas — sem essa checagem, CDbl
' quebraria o macro inteiro na primeira célula em branco que encontrasse.
Private Function NzNum(v As Variant) As Double
    If IsNumeric(v) Then
        NzNum = CDbl(v)
    Else
        NzNum = 0
    End If
End Function

' Cria a aba "CSV" do zero (ou limpa e reaproveita, se já existir de
' uma rodada anterior do macro) e escreve só a linha de cabeçalho — o
' preenchimento de dados é todo feito depois, por EscreverLinhaCSV.
Private Function PrepararAbaCSV(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("CSV")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "CSV"
    Else
        ws.Cells.Clear
    End If

    Dim cabecalhos As Variant, i As Long
    cabecalhos = Array("data", "colaborador", "matricula", "gestor", "setor", "cargo", "tipo_ocorrencia", _
                        "situacao", "status", "duracao_minutos", "destino_horas_extra", "data_tratativa_pontonet", _
                        "horas_excedentes", "ocorrencias_extra_falta_mes_atual", "ocorrencias_extra_falta_3_meses", _
                        "pausas_corretas", "pausas_menor_20min", "pausas_maior_20min", "trabalho_correto_140", _
                        "trabalho_maior_140", "trabalho_menor_140", "pausas_marcacoes_impares")
    For i = LBound(cabecalhos) To UBound(cabecalhos)
        ws.Cells(1, i + 1).Value = cabecalhos(i)
    Next i
    ws.Rows(1).Font.Bold = True

    Set PrepararAbaCSV = ws
End Function

' Ordem das colunas tem que casar com PrepararAbaCSV. Os parâmetros
' opcionais no fim só são preenchidos nas linhas especiais (auditoria de
' extra+falta e resumo de pausas térmicas, uma linha por colaborador);
' nas linhas normais de ocorrência ficam em branco. Use argumentos
' nomeados (ex.: pausaCorretas:=5) pra pular os que não interessam.
Private Sub EscreverLinhaCSV(ws As Worksheet, linha As Long, dataOcorrencia As Variant, nome As String, _
    matricula As String, gestor As String, setorNome As String, cargoNome As String, tipoOcorrencia As String, _
    situacaoTxt As String, statusTxt As String, duracaoMinutos As Double, destinoExtra As String, tratativa As Variant, _
    Optional horasExcedentes As String = "", Optional extraFaltaMesAtual As Variant = Empty, _
    Optional extraFaltaTotal3Meses As Variant = Empty, Optional pausaCorretas As Variant = Empty, _
    Optional pausaMenor20 As Variant = Empty, Optional pausaMaior20 As Variant = Empty, _
    Optional trabalhoCorreto140 As Variant = Empty, Optional trabalhoMaior140 As Variant = Empty, _
    Optional trabalhoMenor140 As Variant = Empty, Optional pausasMarcacoesImpares As Variant = Empty)

    ws.Cells(linha, 1).Value = CDate(dataOcorrencia)
    ws.Cells(linha, 2).Value = nome
    ws.Cells(linha, 3).Value = matricula
    ws.Cells(linha, 4).Value = gestor
    ws.Cells(linha, 5).Value = setorNome
    ws.Cells(linha, 6).Value = cargoNome
    ws.Cells(linha, 7).Value = tipoOcorrencia
    ws.Cells(linha, 8).Value = situacaoTxt
    ws.Cells(linha, 9).Value = statusTxt
    ws.Cells(linha, 10).Value = duracaoMinutos
    ws.Cells(linha, 11).Value = destinoExtra
    If IsDate(tratativa) Then
        ws.Cells(linha, 12).Value = CDate(tratativa)
    Else
        ws.Cells(linha, 12).Value = ""
    End If
    ws.Cells(linha, 13).Value = horasExcedentes
    If Not IsEmpty(extraFaltaMesAtual) Then ws.Cells(linha, 14).Value = extraFaltaMesAtual
    If Not IsEmpty(extraFaltaTotal3Meses) Then ws.Cells(linha, 15).Value = extraFaltaTotal3Meses
    If Not IsEmpty(pausaCorretas) Then ws.Cells(linha, 16).Value = pausaCorretas
    If Not IsEmpty(pausaMenor20) Then ws.Cells(linha, 17).Value = pausaMenor20
    If Not IsEmpty(pausaMaior20) Then ws.Cells(linha, 18).Value = pausaMaior20
    If Not IsEmpty(trabalhoCorreto140) Then ws.Cells(linha, 19).Value = trabalhoCorreto140
    If Not IsEmpty(trabalhoMaior140) Then ws.Cells(linha, 20).Value = trabalhoMaior140
    If Not IsEmpty(trabalhoMenor140) Then ws.Cells(linha, 21).Value = trabalhoMenor140
    If Not IsEmpty(pausasMarcacoesImpares) Then ws.Cells(linha, 22).Value = pausasMarcacoesImpares
End Sub

' Pede ao usuário a pasta de destino dos CSVs exportados por
' ExportarCSVPorGestor. String vazia = usuário cancelou o diálogo; o
' chamador usa isso pra abortar a exportação sem gerar nada.
Private Function EscolherPasta() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker (evita depender da referência "Microsoft Office Object Library")
    fd.Title = "Selecione a pasta onde salvar os arquivos CSV"
    If fd.Show = -1 Then
        EscolherPasta = fd.SelectedItems(1)
    Else
        EscolherPasta = ""
    End If
End Function

' Nome de gestor vira nome de arquivo ("dados_<gestor>.csv"), e nomes
' de gestor podem ter caracteres que o Windows não aceita em nome de
' arquivo (ex.: "Fulano / Substituto"). Troca cada um desses caracteres
' por "_" pra garantir que Workbooks/ADODB.Stream não falhe ao salvar.
Private Function NomeArquivoSeguro(txt As String) As String
    Dim s As String, i As Long
    Dim invalidos As Variant
    s = txt
    invalidos = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
    For i = LBound(invalidos) To UBound(invalidos)
        s = Replace(s, invalidos(i), "_")
    Next i
    NomeArquivoSeguro = s
End Function

' Escreve a aba "CSV" (ou só as linhas de um gestor) num arquivo texto
' em UTF-8, com aspas nos campos que precisarem, para o Painel de
' Ponto (Dashboard-Ponto) conseguir ler nomes com acento sem problema.
Private Sub ExportarLinhasParaCSV(wsCSV As Worksheet, caminho As String, filtroGestor As String)
    Dim ultimaLinha As Long, ultimaColuna As Long
    ultimaLinha = wsCSV.Cells(wsCSV.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = wsCSV.Cells(1, wsCSV.Columns.Count).End(xlToLeft).Column

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open

    Dim c As Long, r As Long
    Dim linhaTxt As String

    linhaTxt = ""
    For c = 1 To ultimaColuna
        If c > 1 Then linhaTxt = linhaTxt & ","
        linhaTxt = linhaTxt & CampoCSV(CStr(wsCSV.Cells(1, c).Value))
    Next c
    stream.WriteText linhaTxt, 1 ' adWriteLine

    Const COL_GESTOR As Long = 4
    For r = 2 To ultimaLinha
        If filtroGestor = "" Or Trim$(CStr(wsCSV.Cells(r, COL_GESTOR).Value)) = filtroGestor Then
            linhaTxt = ""
            For c = 1 To ultimaColuna
                If c > 1 Then linhaTxt = linhaTxt & ","
                linhaTxt = linhaTxt & CampoCSV(TextoCelula(wsCSV.Cells(r, c)))
            Next c
            stream.WriteText linhaTxt, 1
        End If
    Next r

    stream.SaveToFile caminho, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub

' Formata datas como dd/mm/aaaa (ou dd/mm/aaaa hh:mm quando tem hora),
' do jeito que o Painel de Ponto espera.
Private Function TextoCelula(celula As Range) As String
    Dim v As Variant
    v = celula.Value
    If IsDate(v) Then
        If CDate(v) = Int(CDate(v)) Then
            TextoCelula = Format(v, "dd/mm/yyyy")
        Else
            TextoCelula = Format(v, "dd/mm/yyyy hh:mm")
        End If
    Else
        TextoCelula = CStr(v)
    End If
End Function

' Regra padrão de CSV (RFC 4180): um campo só precisa de aspas quando
' contém vírgula, aspas ou quebra de linha — texto "normal" (nomes,
' datas já formatadas, números) sai sem aspas, deixando o arquivo mais
' limpo. Quando precisa de aspas, cada aspas interna vira "" (dobrada),
' senão o CSV ficaria malformado no meio do campo.
Private Function CampoCSV(valor As String) As String
    If InStr(valor, ",") > 0 Or InStr(valor, Chr(34)) > 0 Or InStr(valor, vbLf) > 0 Then
        CampoCSV = Chr(34) & Replace(valor, Chr(34), Chr(34) & Chr(34)) & Chr(34)
    Else
        CampoCSV = valor
    End If
End Function
