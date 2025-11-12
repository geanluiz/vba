Sub Worksheet_Change(ByVal Target As Range)

    Dim ws As Worksheet: Set ws = Worksheets("Cadastro")
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")

    Application.EnableEvents = False
        
    If Not Intersect(Target, dadosCliente.DataBodyRange.Rows(1)) Is Nothing Then 
        Call DesbloquearPlanilha
        Call FormatarCabecalho
        Call BloquearPlanilha
    End If

    Application.EnableEvents = True

End Sub