Attribute VB_Name = "Infra"
Option Explicit

Dim current As String

Sub InserirLinha(Description As String, Valor As Single)

    If ActiveSheet.ProtectContents Then Call DesbloquearPlanilha
    Application.EnableEvents = False

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

    Call FormatarTotais

    If Not ActiveSheet.ProtectContents Then Call BloquearPlanilha
    Application.EnableEvents = True
    
End Sub

Sub ExportarPDF()

    Call DesbloquearPlanilha

    Dim ws As Worksheet: Set ws = Worksheets("Cadastro")
    Dim Tbl As ListObject: Set Tbl = ThisWorkbook.ActiveSheet.ListObjects("OrcamentTbl")
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

    Application.ScreenUpdating = False

    Dim mainWS As Worksheet: Set mainWS = Worksheets("ORÇAMENTO")
    Dim Tbl As ListObject: Set Tbl = mainWS.ListObjects("OrcamentTbl")

    Dim cadastroWS As Worksheet: Set cadastroWS = Worksheets("Cadastro")
    Dim dadosCliente As ListObject: Set dadosCliente = cadastroWS.ListObjects("DadosOrcto")
    Dim vChapas As ListObject: Set vChapas = cadastroWS.ListObjects("ValoresChapas")
    Dim vAcess As ListObject: Set vAcess = cadastroWS.ListObjects("ValoresAcess")
    Dim vGranito As ListObject: Set vGranito = cadastroWS.ListObjects("coresGranito")
    Dim vCubas As ListObject: Set vCubas = cadastroWS.ListObjects("modelosCubas")

    cadastroWS.Unprotect Password:="l123l456"
    mainWS.Unprotect Password:="l123l456"


    mainWS.Range("E3").MergeArea.Locked = True
    mainWS.Range("E4").MergeArea.Locked = True
    mainWS.Range("E5").MergeArea.Locked = True

    ' tabela principal
    If Not Tbl.DataBodyRange Is Nothing Then
        With Tbl
            .DataBodyRange.Columns(1).Locked = True      ' coluna item
            .DataBodyRange.Columns(2).Locked = False      ' coluna descricao
            .DataBodyRange.Columns(3).Locked = False      ' coluna quantidade
            .DataBodyRange.Columns(4).Locked = False      ' coluna valor unit
            .DataBodyRange.Columns(5).Locked = True      ' coluna subtotal
            .TotalsRowRange.Locked = True
        End With
    End If

    ' tabelas auxiliares
    vAcess.HeaderRowRange.Locked = True              ' valores dos acessorios
    dadosCliente.HeaderRowRange.Locked = True        ' dados do cliente/orcamento
    vChapas.DataBodyRange.Rows(1).Locked = False     ' valores das chapas
    vChapas.DataBodyRange.Cells(1, 1).Locked = True  ' texto "chapa"
    vGranito.DataBodyRange.Columns(1).Locked = True
    ' vCubas.HeaderRowRange.Locked = True
    ' vCubas.HeaderRowRange.Columns(1).Locked = True
    
    If mainWS.Range("G1") = "a" Or mainWS.Range("G1") = "t" Then 
        cadastroWS.Visible = xlSheetVisible
    Else 
        cadastroWS.Visible = xlSheetVeryHidden
    End If

    mainWS.Range("G1").Locked = False              ' referencia usada no if abaixo

    
    If mainWS.Range("G1") <> "t" Then 
        cadastroWS.Protect Password:="l123l456"
        mainWS.Protect Password:="l123l456"
    End If

    Application.ScreenUpdating = True
    
End Sub

Sub DesbloquearPlanilha()

    If ActiveSheet.ProtectContents Then 
        Worksheets("Cadastro").Unprotect Password:="l123l456"
        Worksheets("Cadastro").Unprotect
        Worksheets("ORÇAMENTO").Unprotect Password:="l123l456"
        Worksheets("ORÇAMENTO").Unprotect
    End If

End Sub

Sub CallUserForm()

    UserForm1.Show
    
End Sub

Sub CallUserForm2()

    UserForm2.Show
    
End Sub

Sub CallUserForm3()

    UserForm3.Show
    
End Sub

Sub ExcluirLinha()

    Dim ws As Worksheet: Set ws = Worksheets("ORÇAMENTO")
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("OrcamentTbl")
    Dim items As Variant
    Dim inputArray
    Dim i As Integer

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Call DesbloquearPlanilha

    
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

    Call FormatarTabela

    Call FormatarTotais

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Call BloquearPlanilha

End Sub

Sub NovoOrcamento()

    If ActiveSheet.ProtectContents Then Call DesbloquearPlanilha
    Application.EnableEvents = False

    Dim ws As Worksheet: Set ws = Worksheets("Cadastro")
    Dim Tbl As ListObject: Set Tbl = ThisWorkbook.ActiveSheet.ListObjects("OrcamentTbl")
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

    If Not ActiveSheet.ProtectContents Then Call BloquearPlanilha
    Application.EnableEvents = True

End Sub