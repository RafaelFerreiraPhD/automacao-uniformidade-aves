Attribute VB_Name = "GestaoLotesEPesagens"
Option Explicit

' ====================================================================
' REGISTO DE NOVO LOTE E VINCULAÇÃO COM O MANUAL
' ====================================================================
Sub SalvarNovoLote()
    Dim wsNovo As Worksheet, wsCad As Worksheet, wsColeta As Worksheet, wsPainel As Worksheet
    Dim proxLinha As Long
    Dim idLote As String, categoria As String, tipoAve As String, linhagem As String
    Dim dataAlojamento As Variant, numAves As Variant
    Dim ws As Worksheet
    
    On Error GoTo TrataErro
    
    Set wsNovo = Sheets("NOVO LOTE")
    Set wsCad = Sheets("Cadastro_Lote")
    Set wsColeta = Sheets("NOVA COLETA")
    Set wsPainel = Sheets("PAINEL DE CONTROLE")
    
    ' 1. Captura de dados do formulário
    idLote = Trim(wsNovo.Range("E6").Value)
    categoria = Trim(wsNovo.Range("E7").Value)
    tipoAve = Trim(wsNovo.Range("E8").Value)
    linhagem = Trim(wsNovo.Range("E9").Value)
    dataAlojamento = wsNovo.Range("E10").Value
    numAves = wsNovo.Range("E11").Value
    
    ' 2. Validações zootécnicas e operacionais
    If idLote = "" Then
        MsgBox "Por favor, informe o ID DO LOTE!", vbExclamation, "Campo Obrigatório"
        wsNovo.Range("E6").Select
        Exit Sub
    End If
    
    If categoria = "" Or linhagem = "" Then
        MsgBox "Por favor, preencha Categoria e Linhagem!", vbExclamation, "Campo Obrigatório"
        Exit Sub
    End If
    
    If dataAlojamento = "" Or Not IsDate(dataAlojamento) Then
        MsgBox "Por favor, informe uma Data de Alojamento válida!", vbExclamation, "Data Inválida"
        Exit Sub
    End If
    
    If Not IsNumeric(numAves) Or Val(numAves) <= 0 Then
        MsgBox "Por favor, informe um Número de Aves válido!", vbExclamation, "Quantidade Inválida"
        Exit Sub
    End If
    
    ' 3. Gravação sequencial na base Cadastro_Lote
    proxLinha = 2
    Do While Trim(wsCad.Cells(proxLinha, 1).Value) <> ""
        proxLinha = proxLinha + 1
    Loop
    
    wsCad.Cells(proxLinha, 1).Value = idLote
    wsCad.Cells(proxLinha, 2).Value = categoria
    wsCad.Cells(proxLinha, 3).Value = tipoAve
    wsCad.Cells(proxLinha, 4).Value = linhagem
    wsCad.Cells(proxLinha, 5).Value = CDate(dataAlojamento)
    wsCad.Cells(proxLinha, 6).Value = CLng(numAves)
    
    wsCad.Cells(proxLinha, 5).NumberFormat = "dd/mm/yyyy"
    wsCad.Cells(proxLinha, 6).NumberFormat = "#,##0"
    
    ' 4. Limpeza do formulário de entrada
    wsNovo.Range("E6:H11").ClearContents
    
    ' 5. Preparação da recolha e do seletor do painel
    wsColeta.Visible = xlSheetVisible
    wsColeta.Range("C3:D3").Value = idLote
    wsColeta.Range("C4:D4").Value = ""
    wsColeta.Range("E3:E1000").ClearContents
    
    wsPainel.Unprotect
    wsPainel.Range("G3").Value = idLote
    wsPainel.Protect
    
    ' 6. Navegação e transição de ecrã
    Application.ScreenUpdating = False
    wsColeta.Activate
    Application.Goto wsColeta.Range("C4"), True
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "NOVA COLETA" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    Application.ScreenUpdating = True
    
    MsgBox "Lote '" & idLote & "' cadastrado com sucesso! Ecrã pronto para inserção de pesos.", vbInformation, "Registo Efetuado"
    Exit Sub

TrataErro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao salvar o lote: " & Err.Description, vbCritical, "Erro de Execução"
End Sub

' ====================================================================
' VALIDAÇÃO DA PESAGEM, GRAVAÇÃO NO HISTÓRICO E ATUALIZAÇÃO DO PAINEL
' ====================================================================
Sub ChecarResultadosESalvar()
    Dim wsColeta As Worksheet, wsPainel As Worksheet, wsHist As Worksheet
    Dim proxLinha As Long, qtdAmostras As Long
    Dim loteAtual As String, faseCriacao As String, statusUnif As String
    Dim dataColeta As Variant, idadeDias As Variant
    Dim pesoMedio As Double, uniformidade As Double, desvPad As Double
    Dim ws As Worksheet
    
    On Error GoTo TrataErro
    
    Set wsColeta = Sheets("NOVA COLETA")
    Set wsPainel = Sheets("PAINEL DE CONTROLE")
    Set wsHist = Sheets("HISTÓRICO GERAL")
    
    loteAtual = Trim(wsColeta.Range("C3").Value)
    dataColeta = wsColeta.Range("C4").Value
    
    If loteAtual = "" Or loteAtual = "-" Then
        MsgBox "Selecione o Lote na célula C3 antes de salvar!", vbExclamation, "Validação"
        Exit Sub
    End If
    
    If dataColeta = "" Or Not IsDate(dataColeta) Then
        MsgBox "Data de recolha inválida na célula C4!", vbExclamation, "Validação"
        Exit Sub
    End If
    
    qtdAmostras = Application.WorksheetFunction.Count(wsColeta.Range("E3:E1000"))
    If qtdAmostras = 0 Then
        MsgBox "Nenhum peso individual foi inserido na coluna E!", vbExclamation, "Amostra Vazia"
        Exit Sub
    End If
    
    idadeDias = wsColeta.Range("C7").Value
    faseCriacao = wsColeta.Range("C8").Value
    pesoMedio = wsPainel.Range("C5").Value
    desvPad = wsPainel.Range("C6").Value
    uniformidade = wsPainel.Range("C8").Value
    statusUnif = wsPainel.Range("C12").Value
    
    wsHist.Visible = xlSheetVisible
    proxLinha = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
    
    wsHist.Cells(proxLinha, 1).Value = loteAtual
    wsHist.Cells(proxLinha, 2).Value = CDate(dataColeta)
    wsHist.Cells(proxLinha, 3).Value = idadeDias
    wsHist.Cells(proxLinha, 4).Value = faseCriacao
    wsHist.Cells(proxLinha, 5).Value = pesoMedio
    wsHist.Cells(proxLinha, 6).Value = uniformidade
    wsHist.Cells(proxLinha, 7).Value = desvPad
    wsHist.Cells(proxLinha, 8).Value = qtdAmostras
    wsHist.Cells(proxLinha, 9).Value = statusUnif
    
    wsHist.Cells(proxLinha, 2).NumberFormat = "dd/mm/yyyy"
    wsHist.Cells(proxLinha, 5).NumberFormat = "#,##0.0"
    wsHist.Cells(proxLinha, 6).NumberFormat = "0.0%"
    wsHist.Cells(proxLinha, 7).NumberFormat = "#,##0.0"
    wsHist.Cells(proxLinha, 8).NumberFormat = "#,##0"
    
    wsColeta.Range("C3:D4").ClearContents
    wsColeta.Range("E3:E1000").ClearContents
    wsColeta.Calculate
    
    wsPainel.Visible = xlSheetVisible
    wsPainel.Unprotect
    wsPainel.Range("G3").Value = loteAtual
    wsPainel.Protect
    
    Application.ScreenUpdating = False
    wsPainel.Activate
    Application.Goto wsPainel.Range("A1"), True
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "PAINEL DE CONTROLE" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    Application.ScreenUpdating = True
    
    MsgBox "Resultados gravados no Histórico Geral com sucesso!", vbInformation, "Dados Consolidados"
    Exit Sub

TrataErro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao gravar a recolha: " & Err.Description, vbCritical, "Erro de Execução"
End Sub
