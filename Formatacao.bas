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
Dim largT As Single
Dim profT As Single
Dim cuba As Integer
Dim medRod As Single

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

Function desc_tampo(cor, largT, profT, cuba, medRod)

    Dim modelo As String
    Dim sep As String

    
    medRod = medRod * 100

    sep = ", "

    Select Case cuba
        Case "cuba_red":
            modelo = " e cuba Redonda em louça branca."
        Case "cuba_red_slim":
            modelo = " e cuba Redonda Slim em louça branca."
        Case "cuba_ret":
            modelo = " e cuba Retangular em louça branca."
        Case "cuba_ret_slim":
            modelo = " e cuba Retangular Slim em louça branca."
        Case "OptionSemCuba":
            modelo = ". Cuba não inclusa."
            sep = " e "
    End Select

    desc_tampo = "BANHEIRO: Tampo em granito " & cor & " med. " & largT * 100 & "x" & profT * 100 & _
        "cm com rodopia de " & medRod & "cm" & sep & "acabamento em meia esquadria" & modelo

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

    
    If oDate = "" Then
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
    Dim wordLimit as Integer


    'formatar linhas da tabela
    Tbl.DataBodyRange.EntireRow.AutoFit

    'Atualizar index dos itens
    For Each i In Tbl.DataBodyRange.Rows
        If Not i.Cells(1) = i.row - 9 Then
            i.Cells(1) = i.row - 9
        End If
        
        i.Font.Bold = False

        wordLimit = InStr(i.Cells(2).Value, ":") + 1
        i.Cells(2).Characters(Start:=0, Length:=wordLimit).Font.Bold = True

    Next


    Call BloquearPlanilha

End Sub

Sub FormatarTotais()

    Call DesbloquearPlanilha
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim total1 As Range
    Dim total2 As Range

    Set total1 = Tbl.TotalsRowRange.Range(Cells(2), Cells(4))
    Set total2 = Tbl.TotalsRowRange.Cells(5)

    With total1.Borders(xlEdgeLeft)
        .ThemeColor = 1
    End With
    With total1.Borders(xlInsideVertical)
        .ThemeColor = 1
    End With


    Tbl.TotalsRowRange.RowHeight = 35
    

    total1.Font.Bold = True
    total2.Font.Bold = True
    total1.HorizontalAlignment = xlCenter


    Call BloquearPlanilha

End Sub