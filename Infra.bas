Attribute VB_Name = "Infra"
Option Explicit

Dim current As String

Sub InserirLinha(Description As String, Valor As Single)

    Call DesbloquearPlanilha

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim NewRow As Range
    
    Const desc As Integer = 2
    Const Qtde As Integer = 3
    Const Unit As Integer = 4
    Const SubT As Integer = 5
        
    Set NewRow = Tbl.ListRows.Add.Range
    
    If Tbl.DataBodyRange.Cells(1, 2) = "" Then
        Tbl.DataBodyRange.Rows(1).Delete
    End If
        
    With NewRow
        .Cells(desc) = Description
        .Cells(Qtde) = "1"
        .Cells(Unit) = Valor
        .Cells(SubT).Formula = "=[VALOR UNT.]*[QTDE]"
        .EntireRow.AutoFit
    End With

    Call FormatarTabela
    
    Call BloquearPlanilha
    
End Sub

Sub ExportarPDF()

    Call DesbloquearPlanilha

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")

    Dim cName As String
    Dim OrcNum As Integer
    Dim oDate As String
    Dim oPath As String
    
    If Tbl.DataBodyRange Is Nothing Then
        MsgBox "Orçamento sem itens!", "Erro!", vbExclamation
        Exit Sub
    End If

    Call FormatarCabecalho

    cName = dadosCliente.DataBodyRange.Columns(1).Value2
    oDate = dadosCliente.DataBodyRange.Columns(2).Value2
    OrcNum = dadosCliente.DataBodyRange.Columns(3).Value2
    
    oPath = Format(oDate, "yyyy-mm-dd") & " " & OrcNum & " " & UCase(cName)
    oPath = ActiveWorkbook.Path & Application.PathSeparator & oPath & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True, Filename:=oPath
    MsgBox "Salvo em" & oPath, vbInformation, "PDF Salvo"
    
    Call BloquearPlanilha
       
End Sub

Sub BloquearPlanilha()

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")
    Dim vChapas As ListObject: Set vChapas = ws.ListObjects("ValoresChapas")
    Dim vAcess As ListObject: Set vAcess = ws.ListObjects("ValoresAcess")
    Dim vGranito As ListObject: Set vGranito = ws.ListObjects("coresGranito")


    ws.Unprotect Password:="a123a456"


    ' tabela principal
    If Not Tbl.DataBodyRange Is Nothing Then
        With Tbl.DataBodyRange
            .Columns(1).Locked = True       ' coluna de indices tabela principal
            .Columns(2).Locked = False      ' coluna descrição
            .Columns(3).Locked = False      ' coluna quantidade
            .Columns(4).Locked = False      ' coluna valor unit
            .Columns(5).Locked = True       ' coluna valor final do item
        End With
    End If
    Tbl.TotalsRowRange.EntireRow.Locked = True       ' linha de totais

    ' tabelas auxiliares
    vAcess.DataBodyRange.Locked = False              ' valores dos acessórios
    dadosCliente.DataBodyRange.Locked = False        ' dados do cliente/orçamento
    vChapas.DataBodyRange.Rows(1).Locked = False     ' valores das chapas
    vChapas.DataBodyRange.Cells(1, 1).Locked = True  ' texto "chapa"
    vChapas.DataBodyRange.Cells(1, 2).Locked = False ' m² das chapas
    vGranito.DataBodyRange.Columns(1).Locked = True
    

    Range("G1").Locked = False                       ' referência usada no if abaixo

    If Range("G1") <> "t" Then ws.Protect Password:="a123a456"
    
End Sub

Sub DesbloquearPlanilha()

    ActiveSheet.Unprotect Password:="a123a456"
    ActiveSheet.Unprotect

End Sub

Sub CallUserForm()

    UserForm1.Show
    
End Sub

Sub CallUserForm2()

    UserForm2.Show
    
End Sub

Sub ExcluirLinha()

    Call DesbloquearPlanilha

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim items As Variant
    Dim inputArray
    Dim i As Integer

    
    items = Application.InputBox("Qual(is) item(s) deseja excluir? (e.g., 3 ou 2-5)", "Excluir Linhas", Type:=2)
    If items = False Then
        Exit Sub
    End If
    
    inputArray = Split(items, "-")
    ' Cast items inside inputArray from [sub]string to integer
    For i = LBound(inputArray) To UBound(inputArray)
        inputArray(i) = CInt(inputArray(i))
    Next i


    If UBound(inputArray) > 1 Then 
        MsgBox "Erro! Verifique as informações digitadas e tente novamente...", "Erro!", vbExclamation
    ElseIf UBound(inputArray) = 0 Then
        Tbl.ListRows(inputArray(0)).Delete
    Else
        For i = inputArray(0) To inputArray(1)
            Tbl.ListRows(inputArray(0)).Delete
        Next i
    End If

    
    If Tbl.ListRows.Count = 0 Then Tbl.ListRows.Add


    Call FormatarTotais

    Call FormatarTabela

    Call BloquearPlanilha

End Sub

Sub NovoOrcamento()

    Call DesbloquearPlanilha

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")

    Dim rowCount As Integer
    Dim NewRow As Range

    rowCount = Tbl.ListRows.Count

    While rowCount > 0
        Tbl.DataBodyRange.Rows(rowCount).Delete
        rowCount = Tbl.ListRows.Count
    Wend

    Set NewRow = Tbl.ListRows.Add.Range

    dadosCliente.DataBodyRange.Columns(1).Value2 = ""
    dadosCliente.DataBodyRange.Columns(2).Value2 = ""
    dadosCliente.DataBodyRange.Columns(3).Value2 = dadosCliente.DataBodyRange.Columns(3).Value2 + 1

    Call FormatarTotais
    Call FormatarCabecalho

    Call BloquearPlanilha

End Sub

Sub MoverMenu()

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim btns As Shape: Set btns = ws.Shapes("GrupoBtns")

    If Range("H2") = "" Then
        btns.Top = 25
    Else
        btns.Top = 70
    End If

    btns.Left = 640

End Sub

Sub MostrarTabela(table As String)
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")
    Dim valoresChapas As ListObject: Set valoresChapas = ws.ListObjects("ValoresChapas")
    Dim valoresAcess As ListObject: Set valoresAcess = ws.ListObjects("ValoresAcess")
    Dim valoresGranito As ListObject: Set valoresGranito = ws.ListObjects("coresGranito")
    Dim valoresCubas As ListObject: Set valoresCubas = ws.ListObjects("modelosCubas")


    Select Case table
        Case "Cliente":
            If dadosCliente.Range.Address = "$H$2:$L$3" Then
                dadosCliente.Range.Cut Range("$U$8:$Y$9")
            Else
                dadosCliente.Range.Cut Range("$H$2:$L$3")
                current = table
            End If
        Case "Chapas":
            If valoresChapas.Range.Address = "$H$2:$M$3" Then
                valoresChapas.Range.Cut Range("$U$2:$Z$3")
            Else
                valoresChapas.Range.Cut Range("$H$2:$M$3")
                current = table
            End If
        Case "Acess":
            If valoresAcess.Range.Address = "$H$2:$M$3" Then
                valoresAcess.Range.Cut Range("$U$5:$Z$6")
            Else
                valoresAcess.Range.Cut Range("$H$2:$M$3")
                current = table
            End If
        Case "Granito":
            If valoresGranito.Range.Address = "$H$2:$L$3" Then
                valoresGranito.Range.Cut Range("$AB$2:$AF$3")
            Else
                valoresGranito.Range.Cut Range("$H$2:$L$3")
                current = table
            End If
        Case "Cubas":
            If valoresCubas.Range.Address = "$H$2:$L$3" Then
                valoresCubas.Range.Cut Range("$AB$5:$AF$6")
            Else
                valoresCubas.Range.Cut Range("$H$2:$L$3")
                current = table
            End If
    End Select

End Sub

Sub MenuClick()

    Call DesbloquearPlanilha

    If Range("$H$2").Value2 = "" Or current = Application.Caller() Then
        Call MostrarTabela(Application.Caller())
    Else
        Call MostrarTabela(current)
        Call MostrarTabela(Application.Caller())
    End If

    MoverMenu

    Call BloquearPlanilha

End Sub