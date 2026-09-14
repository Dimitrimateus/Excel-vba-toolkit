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
'     dia": coluna "BH" para as horas de falta cobertas por banco de
'     horas, e a soma de "50%"+"60%"+"100%"+"120%" para hora extra.
'     Se não achar a linha correspondente no Cartão Ponto, cai de
'     volta para o valor da aba Tratamento (para não perder dado).
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
    Set wsCartao = EncontrarAba(wbTrat, "Cartão ponto")

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
        If UCase$(checkTxt) = VALOR_CHECK_IGNORAR Then GoTo ProximaLinha
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

        ' Quantidade de horas "oficial": Cartão Ponto (BH = falta coberta
        ' por banco de horas; 50%+60%+100%+120% = hora extra). Cai de
        ' volta para o valor da aba Tratamento se não achar a linha
        ' correspondente no Cartão Ponto (matrícula + data).
        Dim minFaltaBHCartao As Double, minExtraCartao As Double
        ObterValoresCartao dictCartao, matricula, dataOcorrencia, minFaltaBHCartao, minExtraCartao

        Dim minExtra As Double, minFalta As Double
        minExtra = IIf(minExtraCartao > 0, minExtraCartao, minExtraTrat + minExtra100Trat)
        minFalta = IIf(minFaltaBHCartao > 0, minFaltaBHCartao, minFaltaTrat)

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

        ' Linha de HORA EXTRA, se a aba Tratamento apontar extra nesse dia
        ' (gate). O valor em minutos vem do Cartão Ponto (ou, na falta
        ' dele, da própria Tratamento). Não recebe a demora do PontoNet:
        ' essa demora é sobre a justificativa de ausência de marcação,
        ' não sobre a extra.
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

    If Not wbAus Is Nothing Then wbAus.Close SaveChanges:=False

    MsgBox (linhaSaida - 2) & " linha(s) geradas na aba CSV.", vbInformation, "Gerar CSV"
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
' "BH" (falta coberta por banco de horas) e a soma de
' "50%"+"60%"+"100%"+"120%" (hora extra). Se a aba não existir (não
' foi encontrada na planilha), devolve um dicionário vazio e o
' chamador cai de volta para os valores da aba Tratamento.
Private Function MontarDicionarioCartao(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set MontarDicionarioCartao = dict
        Exit Function
    End If

    Dim colMat As Long, colData As Long, colBH As Long
    Dim col50 As Long, col60 As Long, col100 As Long, col120 As Long
    colMat = ColunaPorCabecalho(ws, "Matricula")
    colData = ColunaPorCabecalho(ws, "DT")
    colBH = ColunaPorCabecalho(ws, "BH")
    col50 = ColunaPorCabecalho(ws, "50%")
    col60 = ColunaPorCabecalho(ws, "60%")
    col100 = ColunaPorCabecalho(ws, "100%")
    col120 = ColunaPorCabecalho(ws, "120%")
    If colMat = 0 Or colData = 0 Then
        Set MontarDicionarioCartao = dict
        Exit Function
    End If

    Dim ultimaLinha As Long, r As Long
    Dim mat As String, dt As Variant, chave As String
    Dim bhMin As Double, extraMin As Double

    ultimaLinha = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row
    For r = 2 To ultimaLinha
        mat = Trim$(CStr(ws.Cells(r, colMat).Value))
        dt = ws.Cells(r, colData).Value
        If mat <> "" And IsDate(dt) Then
            chave = mat & "|" & Format(CDate(dt), "yyyy-mm-dd")
            bhMin = IIf(colBH > 0, NzNum(ws.Cells(r, colBH).Value), 0) * 24 * 60
            extraMin = 0
            If col50 > 0 Then extraMin = extraMin + NzNum(ws.Cells(r, col50).Value) * 24 * 60
            If col60 > 0 Then extraMin = extraMin + NzNum(ws.Cells(r, col60).Value) * 24 * 60
            If col100 > 0 Then extraMin = extraMin + NzNum(ws.Cells(r, col100).Value) * 24 * 60
            If col120 > 0 Then extraMin = extraMin + NzNum(ws.Cells(r, col120).Value) * 24 * 60
            ' se a matrícula tiver mais de uma linha no mesmo dia, soma
            If dict.Exists(chave) Then
                Dim antigo As Variant
                antigo = dict(chave)
                dict(chave) = Array(antigo(0) + bhMin, antigo(1) + extraMin)
            Else
                dict(chave) = Array(bhMin, extraMin)
            End If
        End If
    Next r
    Set MontarDicionarioCartao = dict
End Function

Private Sub ObterValoresCartao(dictCartao As Object, matricula As String, dataOcorrencia As Variant, _
    ByRef minFaltaBH As Double, ByRef minExtra As Double)

    minFaltaBH = 0
    minExtra = 0
    If dictCartao Is Nothing Then Exit Sub
    If Not IsDate(dataOcorrencia) Then Exit Sub

    Dim chave As String
    chave = matricula & "|" & Format(CDate(dataOcorrencia), "yyyy-mm-dd")
    If dictCartao.Exists(chave) Then
        Dim arr As Variant
        arr = dictCartao(chave)
        minFaltaBH = arr(0)
        minExtra = arr(1)
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
                        "horas_excedentes")
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
    Optional horasExcedentes As String = "")

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
