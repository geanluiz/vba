Option Explicit

Sub Worksheet_Open()

    Application.EnableEvents = True

    Dim ws As Worksheet: Set ws = Worksheets("ORÇAMENTO")
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")
    
    dadosCliente.DataBodyRange.Columns(1).Value = ""
    dadosCliente.DataBodyRange.Columns(2).Value = ""

End Sub

Sub Worksheet_Change(ByVal Target As Range)

    Application.EnableEvents = False

    Call DesbloquearPlanilha
    
    If Not Intersect(Target, Range("G1")) Is Nothing Or Range("G1") = "" Then 
        Call BloquearPlanilha
    End If
    
    Application.EnableEvents = True
End Sub