Attribute VB_Name = "NavegacaoSistema"
Option Explicit

Sub NavegarPara(nomeAbaDestino As String)
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    
    With Sheets(nomeAbaDestino)
        .Visible = xlSheetVisible
        .Activate
        .Range("A1").Select
    End With
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> nomeAbaDestino Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    Application.ScreenUpdating = True
End Sub

Sub IrParaNovoLote()
    NavegarPara "NOVO LOTE"
End Sub

Sub IrParaNovaColeta()
    NavegarPara "NOVA COLETA"
End Sub

Sub IrParaPainelControle()
    NavegarPara "PAINEL DE CONTROLE"
End Sub

Sub IrParaHistoricoGeral()
    NavegarPara "HISTÓRICO GERAL"
End Sub

Sub IrParaConfiguracoes()
    NavegarPara "CONFIGURAÇÕES"
End Sub

Sub VoltarAoInicio()
    NavegarPara "INICIO"
End Sub

Sub IrParaCadastroLotes()
    NavegarPara "Cadastro_Lote"
End Sub

Sub IrParaManualLinhagens()
    NavegarPara "Manual_Linhagens"
End Sub

Sub IrParaSuporte_Linhagens()
    NavegarPara "Suporte_Linhagens"
End Sub

Sub VoltarParaConfiguracoes()
    NavegarPara "CONFIGURAÇÕES"
End Sub
