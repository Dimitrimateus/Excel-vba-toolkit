Attribute VB_Name = "GerarCSVPonto"
Option Explicit

' =====================================================================
' GerarCSVPonto
' ---------------------------------------------------------------------
' Consolida a aba "Tratamento" (desta planilha) com os dados de
' "Ausência de marcação" do PontoNet numa aba "CSV", no formato
' esperado pelo Painel de Ponto (Dashboard-Ponto).
'
' Este módulo NÃO foi testado dentro do Excel/VBA (o ambiente onde ele
' foi escrito não tem Excel instalado). Revise com atenção e rode
' primeiro numa CÓPIA da planilha antes de usar com os dados reais.
'
' Como usar:
'   1. Abra a planilha "Tratamento Ponto" (que contém a aba
'      "Tratamento" e a aba "RE 08.09").
'   2. Cole os dados de "Ausência de marcação" numa aba chamada
'      "PontoNet", dentro dessa mesma planilha (mesmo cabeçalho do
'      relatório original: Data falta, Matrícula, Nome, Situação
'      atual, Justificativa, Avaliado em:, etc.).
'   3. Importe este módulo (Alt+F11 > Arquivo > Importar Arquivo).
'   4. Rode a macro GerarAbaCSV. Ela procura a aba "PontoNet"
'      automaticamente; se não encontrar, pede para você selecionar
'      um arquivo separado (pode cancelar essa caixa também; a demora
'      no PontoNet fica em branco nesse caso, o resto continua
'      funcionando).
'   5. Rode a macro ExportarCSVPorGestor. Ela pede uma pasta e cria
'      um arquivo .csv por gestor, mais um "dados_TODOS.csv" com tudo
'      (esse último é o que vai para o supervisor geral).
'
' Regras já confirmadas com o RH:
'   - Linhas com Situação = "Sem alteração" são excluídas (não são
'     ocorrência real).
'   - Linhas com a coluna "Check" = "S" também são excluídas: "S"
'     marca casos que o sistema aponta como ocorrência mas que na
'     prática não têm problema (ex.: 1 minuto de hora extra). Só as
'     marcadas "N" entram no CSV.
'   - A coluna "Cargo" já existe na aba Tratamento (adicionada pelo
'     RH) e é usada direto; só cai no lookup pela "RE 08.09" quando
'     vier vazia.
'   - A coluna "Pontonet" da aba Tratamento é ignorada de propósito
'     (não representa quantidade de horas, é só um status auxiliar).
'   - As colunas "Extras"/"Faltas" da aba Tratamento decidem QUAIS
'     dias entram como ocorrência de hora extra / falta (e são a
'     única fonte para "extra e falta no mesmo dia"). Mas a
'     QUANTIDADE de horas gravada no CSV vem da aba "Cartão ponto até
'     dia": coluna "BH" (coluna P) para as horas de falta cobertas por
'     banco de horas; coluna "50%" (coluna S) para hora extra "normal".
'     Se não achar a linha correspondente no Cartão Ponto, cai de
'     volta para o valor da aba Tratamento (para não perder dado).
'   - Hora extra a 100% (coluna "100%", coluna U, do Cartão Ponto) é
'     reportada em linha separada, com tipo_ocorrencia = "Hora Extra
'     100%", sempre que existir (Check "S" ou "N") — o painel mostra
'     essa informação no card de "Total de horas extra". As colunas
'     "60%" e "120%" NÃO são usadas.
'   - Ocorrências curtas (< 15min): quando uma linha tem Check = "S"
'     (que normalmente seria descartada por inteiro) E o Cartão Ponto
'     confirma "Falta < 15min" ou "Extra < 15min" (coluna "Tolerância
'     < 15min", coluna AA), essa linha entra no CSV mesmo assim, só
'     pra alimentar a visão "Ocorrências curtas" do painel — sem
'     contar nos totais normais de hora extra/falta.
'   - "Horas Excedentes" (aba Tratamento, coluna "SIM"/vazio) vira uma
'     lista separada no painel ("Relação de horas excedentes"), com
'     as ocorrências marcadas "SIM".
'
' Pontos que ainda dependem de uma regra de negócio que eu não
' consegui confirmar só olhando os dados (procure "REGRA:" abaixo):
'   - Como decidir se uma hora extra foi para BANCO DE HORAS ou para
'     PAGAMENTO (a aba "Tratamento" não tem essa informação separada;
'     a coluna "BH" do Cartão Ponto é usada para FALTA coberta por
'     banco de horas, não para destino da hora extra).
'   - Como mapear os status reais (Tratado/Tratando/Tratar/vazio) para
'     o vocabulário Pendente/Aprovado/Reprovado/Regularizado do painel.
'   - Na aba Cartão Ponto, a coluna "BH" aparece tanto em dias de
'     "Falta (Banco Horas)" (valor alto, o dia inteiro) quanto em dias
'     "Trabalhando" (valor baixo, minutos). Este código trata qualquer
'     valor de "BH" como hora de falta coberta pelo banco, do jeito
'     que foi pedido. Se, na prática, o "BH" de um dia "Trabalhando"
'     for crédito (e não falta), avise para eu ajustar.
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

    Dim dictRE As Object, dictAus As Object, dictCartao As Object, colIdx As Object
    Set dictRE = MontarDicionarioRE(wsRE)
    Set dictAus = MontarDicionarioAusencia(wsAus)
    Set dictCartao = MontarDicionarioCartao(wsCartao)
    Set colIdx = MapearColunas(wsTrat)

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

        If nome = "" Or Not IsDate(dataOcorrencia) Then GoTo ProximaLinha

        If gestor <> "" And matricula <> "" Then
            If Not dictGestorPorMatricula.Exists(matricula) Then dictGestorPorMatricula.Add matricula, gestor
        End If

        ' Cartão Ponto: consultado sempre, mesmo em linhas Check = "S",
        ' porque a hora extra a 100% e as ocorrências curtas (<15min)
        ' são reportadas independente do Check.
        Dim minFaltaBHCartao As Double, minExtra50Cartao As Double, minExtra100Cartao As Double
        Dim tolerancia15 As String
        ObterValoresCartao dictCartao, matricula, dataOcorrencia, minFaltaBHCartao, minExtra50Cartao, minExtra100Cartao, tolerancia15

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
            ' Check = "S": não é uma ocorrência real pro RH, mas se o
            ' Cartão Ponto confirmar que foi uma falta/extra < 15min,
            ' registra só essa linha (pra alimentar "Ocorrências
            ' curtas" no painel) e pula o resto do fluxo normal, sem
            ' contar de novo nos totais de hora extra/falta.
            If tolerancia15 <> "" Then
                Dim valorCurta As Double
                If InStr(1, tolerancia15, "Falta", vbTextCompare) > 0 Then
                    valorCurta = -Round(minFaltaBHCartao, 0)
                Else
                    valorCurta = Round(minExtra50Cartao + minExtra100Cartao, 0)
                End If
                EscreverLinhaCSV wsCSV, linhaSaida, dataOcorrencia, nome, matricula, gestor, setorNome, cargoNome, _
                    tolerancia15, tolerancia15, statusFinal, valorCurta, "", Empty, horasExcedentesTxt
                linhaSaida = linhaSaida + 1
            End If
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

                EscreverLinhaCSV wsCSV, linhaSaida, DataMaximaAba(wsCartao), nomeAud, CStr(matAud), gestorAud, _
                    setorAud, cargoAud, "Auditoria Extra e Falta (3 Meses)", "", "Pendente", 0, "", Empty, "", _
                    mesAtualCount, total3Meses
                linhaSaida = linhaSaida + 1
                qtdColaboradoresAuditoria = qtdColaboradoresAuditoria + 1
            End If
        Next matAud
    End If

    If Not wbAus Is Nothing Then wbAus.Close SaveChanges:=False

    Dim msgAuditoria As String
    If abasCartao.Count > 0 Then
        msgAuditoria = vbCrLf & qtdColaboradoresAuditoria & " colaborador(es) com extra+falta no mesmo dia nos últimos " & _
            abasCartao.Count & " mês(es) (" & nomeAbaAtual & " = mês atual)."
    Else
        msgAuditoria = vbCrLf & "Nenhuma aba ""Cartão ponto ..."" encontrada; auditoria de 3 meses não gerada."
    End If

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

' Acha uma aba cujo nome comece com "prefixo" (ignora maiúsculas e
' espaços extras no fim do nome, tipo "Cartão ponto até dia "),
' Returns Nothing se não encontrar.
Private Function EncontrarAba(wb As Workbook, prefixo As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If LCase$(Left$(Trim$(ws.Name), Len(prefixo))) = LCase$(prefixo) Then
            Set EncontrarAba = ws
            Exit Function
        End If
    Next ws
    Set EncontrarAba = Nothing
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
    colData = ColunaPorCabecalho(ws, "DT")
    If colData = 0 Then
        DataMaximaAba = maxData
        Exit Function
    End If
    ultimaLinha = ws.Cells(ws.Rows.Count, colData).End(xlUp).Row
    For r = 2 To ultimaLinha
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
    colMat = ColunaPorCabecalho(ws, "Matrícula")
    colNome = ColunaPorCabecalho(ws, "Nome")
    colFlag = ColunaPorCabecalho(ws, "Banco de horas e extra no mesmo dia")
    If colMat = 0 Or colFlag = 0 Then
        Set ContarExtraFaltaMesmoDia = dict
        Exit Function
    End If

    Dim ultimaLinha As Long, r As Long, mat As String, flagTxt As String, nome As String
    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 2 To ultimaLinha
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

Private Function ColunaPorCabecalho(ws As Worksheet, cabecalho As String) As Long
    Dim ultimaColuna As Long, c As Long
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultimaColuna
        If Trim$(CStr(ws.Cells(1, c).Value)) = cabecalho Then
            ColunaPorCabecalho = c
            Exit Function
        End If
    Next c
    ColunaPorCabecalho = 0
End Function

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
' extra a 100%) — "60%" e "120%" não entram, por pedido do RH. Guarda
' também o texto da coluna "Tolerância < 15min" ("Falta < 15min" /
' "Extra < 15min"), usado para a visão "Ocorrências curtas". Se a aba
' não existir (não foi encontrada na planilha), devolve um dicionário
' vazio e o chamador cai de volta para os valores da aba Tratamento.
Private Function MontarDicionarioCartao(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set MontarDicionarioCartao = dict
        Exit Function
    End If

    Dim colMat As Long, colData As Long, colBH As Long
    Dim col50 As Long, col100 As Long, colTolerancia As Long
    colMat = ColunaPorCabecalho(ws, "Matricula")
    colData = ColunaPorCabecalho(ws, "DT")
    colBH = ColunaPorCabecalho(ws, "BH")
    col50 = ColunaPorCabecalho(ws, "50%")
    col100 = ColunaPorCabecalho(ws, "100%")
    colTolerancia = ColunaPorCabecalho(ws, "Tolerância < 15min")
    If colMat = 0 Or colData = 0 Then
        Set MontarDicionarioCartao = dict
        Exit Function
    End If

    Dim ultimaLinha As Long, r As Long
    Dim mat As String, dt As Variant, chave As String
    Dim bhMin As Double, extra50Min As Double, extra100Min As Double, tolerTxt As String

    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 2 To ultimaLinha
        mat = Trim$(CStr(ws.Cells(r, colMat).Value))
        dt = ws.Cells(r, colData).Value
        If mat <> "" And IsDate(dt) Then
            chave = mat & "|" & Format(CDate(dt), "yyyy-mm-dd")
            bhMin = IIf(colBH > 0, NzNum(ws.Cells(r, colBH).Value), 0) * 24 * 60
            extra50Min = IIf(col50 > 0, NzNum(ws.Cells(r, col50).Value), 0) * 24 * 60
            extra100Min = IIf(col100 > 0, NzNum(ws.Cells(r, col100).Value), 0) * 24 * 60
            tolerTxt = IIf(colTolerancia > 0, Trim$(CStr(ws.Cells(r, colTolerancia).Value)), "")
            ' se a matrícula tiver mais de uma linha no mesmo dia, soma
            ' os minutos e guarda o último texto de tolerância não vazio
            If dict.Exists(chave) Then
                Dim antigo As Variant
                antigo = dict(chave)
                If tolerTxt = "" Then tolerTxt = antigo(3)
                dict(chave) = Array(antigo(0) + bhMin, antigo(1) + extra50Min, antigo(2) + extra100Min, tolerTxt)
            Else
                dict(chave) = Array(bhMin, extra50Min, extra100Min, tolerTxt)
            End If
        End If
    Next r
    Set MontarDicionarioCartao = dict
End Function

Private Sub ObterValoresCartao(dictCartao As Object, matricula As String, dataOcorrencia As Variant, _
    ByRef minFaltaBH As Double, ByRef minExtra50 As Double, ByRef minExtra100 As Double, ByRef tolerancia15 As String)

    minFaltaBH = 0
    minExtra50 = 0
    minExtra100 = 0
    tolerancia15 = ""
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
        tolerancia15 = arr(3)
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

Private Function NzNum(v As Variant) As Double
    If IsNumeric(v) Then
        NzNum = CDbl(v)
    Else
        NzNum = 0
    End If
End Function

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
                        "horas_excedentes", "ocorrencias_extra_falta_mes_atual", "ocorrencias_extra_falta_3_meses")
    For i = LBound(cabecalhos) To UBound(cabecalhos)
        ws.Cells(1, i + 1).Value = cabecalhos(i)
    Next i
    ws.Rows(1).Font.Bold = True

    Set PrepararAbaCSV = ws
End Function

' Ordem das colunas tem que casar com PrepararAbaCSV.
Private Sub EscreverLinhaCSV(ws As Worksheet, linha As Long, dataOcorrencia As Variant, nome As String, _
    matricula As String, gestor As String, setorNome As String, cargoNome As String, tipoOcorrencia As String, _
    situacaoTxt As String, statusTxt As String, duracaoMinutos As Double, destinoExtra As String, tratativa As Variant, _
    Optional horasExcedentes As String = "", Optional extraFaltaMesAtual As Variant = Empty, _
    Optional extraFaltaTotal3Meses As Variant = Empty)

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
End Sub

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

Private Function CampoCSV(valor As String) As String
    If InStr(valor, ",") > 0 Or InStr(valor, Chr(34)) > 0 Or InStr(valor, vbLf) > 0 Then
        CampoCSV = Chr(34) & Replace(valor, Chr(34), Chr(34) & Chr(34)) & Chr(34)
    Else
        CampoCSV = valor
    End If
End Function
