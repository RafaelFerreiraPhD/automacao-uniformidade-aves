Attribute VB_Name = "ConfiguracoesEAdmin"
Option Explicit

Sub CriarDesignAbaConfiguracoes()
    Dim ws As Worksheet, shp As Shape, btn As Shape
    Dim abaExiste As Boolean
    Dim r As Long
    
    Application.ScreenUpdating = False
    abaExiste = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "CONFIGURAÇÕES" Then
            abaExiste = True
            Set ws = Sheets("CONFIGURAÇÕES")
            Exit For
        End If
    Next ws
    
    If Not abaExiste Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "CONFIGURAÇÕES"
    End If
    
    ws.Cells.Clear
    For Each shp In ws.Shapes: shp.Delete: Next shp
    
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Segoe UI"
    
    ws.Columns("A").ColumnWidth = 4: ws.Columns("B").ColumnWidth = 6: ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 16: ws.Columns("E").ColumnWidth = 6: ws.Columns("F").ColumnWidth = 26
    ws.Columns("G").ColumnWidth = 26: ws.Columns("H").ColumnWidth = 4
    
    With ws.Range("B2:G3")
        .Merge: .Value = "CONFIGURAÇÕES E PARÂMETROS ZOOTÉCNICOS"
        .Font.Bold = True: .Font.Size = 14: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 41, 59): .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    
    With ws.Range("B6:D6")
        .Merge: .Value = "PARÂMETROS DE UNIFORMIDADE"
        .Font.Bold = True: .Font.Size = 10: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(51, 65, 85): .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    
    ws.Range("B7").Value = "•": ws.Range("C7").Value = "Faixa de Tolerância (±% Média)": ws.Range("D7").Value = 0.1: ws.Range("D7").NumberFormat = "0.0%"
    ws.Range("B8").Value = "•": ws.Range("C8").Value = "Uniformidade Excelente (>=)": ws.Range("D8").Value = 0.85: ws.Range("D8").NumberFormat = "0.0%"
    ws.Range("B9").Value = "•": ws.Range("C9").Value = "Uniformidade Boa (>=)": ws.Range("D9").Value = 0.8: ws.Range("D9").NumberFormat = "0.0%"
    ws.Range("B10").Value = "•": ws.Range("C10").Value = "Uniformidade Regular (>=)": ws.Range("D10").Value = 0.7: ws.Range("D10").NumberFormat = "0.0%"
    ws.Range("B11").Value = "•": ws.Range("C11").Value = "Uniformidade Crítica": ws.Range("D11").Value = "< 70.0%": ws.Range("D11").HorizontalAlignment = xlCenter
    
    With ws.Range("B7:D11")
        .Interior.Color = RGB(248, 250, 252): .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(226, 232, 240)
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub ResetarSistema()
    Dim resposta As VbMsgBoxResult, senhaDigitada As String
    Dim tbl As ListObject
    Dim wsCad As Worksheet, wsHist As Worksheet, wsColeta As Worksheet, wsNovo As Worksheet, wsPainel As Worksheet
    Const SENHA_MESTRE As String = "admin123"
    
    resposta = MsgBox("ATENÇÃO: Deseja apagar todos os lotes e pesagens acumuladas?", vbCritical + vbYesNo + vbDefaultButton2, "Confirmação")
    If resposta <> vbYes Then Exit Sub
    
    senhaDigitada = InputBox("Digite a senha de administrador:", "Autenticação")
    If senhaDigitada <> SENHA_MESTRE Then
        MsgBox "Senha incorreta!", vbCritical, "Acesso Negado"
        Exit Sub
    End If
    
    On Error GoTo TrataErro
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    
    Set wsCad = Sheets("Cadastro_Lote"): Set wsHist = Sheets("HISTÓRICO GERAL")
    Set wsColeta = Sheets("NOVA COLETA"): Set wsNovo = Sheets("NOVO LOTE"): Set wsPainel = Sheets("PAINEL DE CONTROLE")
    
    On Error Resume Next
    wsCad.Unprotect: wsHist.Unprotect: wsColeta.Unprotect: wsNovo.Unprotect: wsPainel.Unprotect
    On Error GoTo TrataErro
    
    If wsCad.Cells(wsCad.Rows.Count, "A").End(xlUp).Row > 1 Then wsCad.Range("A2:F" & wsCad.Cells(wsCad.Rows.Count, "A").End(xlUp).Row).ClearContents
    If wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row > 1 Then wsHist.Range("A2:I" & wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row).ClearContents
    
    wsColeta.Range("C3:C4").ClearContents
    wsColeta.Range("E3:E1000").ClearContents
    wsNovo.Range("E6:E11").ClearContents
    wsPainel.Range("G3").Value = ""
    
    Application.DisplayAlerts = True: Application.ScreenUpdating = True
    MsgBox "Dados reinicializados com sucesso!", vbInformation, "Concluído"
    Sheets("INICIO").Activate
    Exit Sub

TrataErro:
    Application.DisplayAlerts = True: Application.ScreenUpdating = True
    MsgBox "Erro: " & Err.Description, vbCritical, "Falha"
End Sub
