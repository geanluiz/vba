Option Explicit

Sub Worksheet_Open()

    Application.EnableEvents = True

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")
    
    dadosCliente.DataBodyRange.Columns(1).Value = ""
    dadosCliente.DataBodyRange.Columns(2).Value = ""

End Sub

Sub Worksheet_Change(ByVal Target As Range)

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")
    
    If Target.Address = dadosCliente.DataBodyRange.Cells(1, 1).Address Then
        Application.EnableEvents = False
        Call FormatarCabecalho
        Application.EnableEvents = True
    End If

End Sub

