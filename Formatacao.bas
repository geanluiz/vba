Attribute VB_Name = "Formatacao"

Option Explicit

Dim cor As String
Dim mold As String
Dim moldRpdAplq As String
Dim largS As Single
Dim largI As Single
Dim altS As Single
Dim altI As Single
Dim profS As Single
Dim profI As Single
Dim medSup As String
Dim medInf As String
Dim qtPInf As Integer
Dim trechoPortas As String

Function desc_text(largS, largI, altS, altI, profS, profI, cor, moldRpdAplq, qtPInf) As String

    If moldRpdAplq = "mold" Then
        mold = " com moldura"
    ElseIf moldRpdAplq = "aplq" Then
        mold = " com moldura e almofada"
    Else
        mold = " ripadas"
    End If

    medSup = (largS * 100) & "x" & (altS * 100) & "x" & (profS * 100)
    medInf = (largI * 100) & "x" & (altI * 100) & "x" & (profI * 100)

    If qtPInf > 1 Then
        trechoPortas = "com " & qtPInf & " portas"
    ElseIf qtPInf = 0 Then
        trechoPortas = "sem portas"
    ElseIf qtPInf = 1 Then
        trechoPortas = "com 1 porta"
    End If


    desc_text = _
    "BANHEIRO: Armário suspenso a prova dagua pintado na cor " & _
    cor & " med. " & medSup & "cm com 1 porta, molduras, espelho e" & _
    " ferragens em inox e; Balcão a prova dagua pintado na cor " & _
    cor & " med. " & medInf & "cm " & trechoPortas & _
    mold & " e ferragens em inox."

End Function

Sub FormatarCabecalho()

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim dadosCliente As ListObject: Set dadosCliente = ws.ListObjects("DadosOrcto")
    
    Dim cName As String
    Dim oDate As Variant
    Dim oNum As Integer


    cName = dadosCliente.DataBodyRange.Columns(1).Value
    oDate = dadosCliente.DataBodyRange.Columns(2).Value
    oNum = dadosCliente.DataBodyRange.Columns(3).Value

    
    If oDate = "" Or Not oDate = Date Then
        oDate = Date
        dadosCliente.DataBodyRange.Cells(2).Value = oDate
    End If

    If cName = "" Then
        cName = UCase(InputBox("Insira o Nome do cliente"))
        dadosCliente.DataBodyRange.Columns(1).Value = cName
    Else
        cName = UCase(cName)
        dadosCliente.DataBodyRange.Columns(1).Value = UCase(cName)
    End If

    Range("E3").Value = "CLIENTE: " & cName
    Range("E4").Value = "DATA: " & oDate
    Range("E5").Value = "ORÇAMENTO Nº " & oNum
    
End Sub

Sub FormatarTabela()

    Call DesbloquearPlanilha

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim i As Range


    'formatar linhas da tabela
    Tbl.DataBodyRange.EntireRow.AutoFit

    'Atualizar index dos itens
    For Each i In Tbl.DataBodyRange.Rows
        If Not i.Cells(1) = i.row - 9 Then
            i.Cells(1) = i.row - 9
        End If
    Next

    Call BloquearPlanilha

End Sub

Sub FormatarTotais()

    Call DesbloquearPlanilha
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")

    Tbl.TotalsRowRange.RowHeight = 35

    Call BloquearPlanilha

End Sub