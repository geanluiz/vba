VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Dados do orçamento"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()

    Dim ws As Worksheet: Set ws = Worksheets("Cadastro")
    Dim dadosOrcto As ListObject: Set dadosOrcto = ws.ListObjects("DadosOrcto")

    Dim i As Integer

    Txt_Cliente = dadosOrcto.DataBodyRange.Cells(1, 1)
    Txt_Data = dadosOrcto.DataBodyRange.Cells(1, 2)
    Txt_Orcto = dadosOrcto.DataBodyRange.Cells(1, 3)


    For i = 1 To 5
        Cmb_Tabela.AddItem i
    Next i

    Cmb_Tabela = dadosOrcto.DataBodyRange.Cells(1, 4)

End Sub

Private Sub Txt_Data_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim enteredDate As String
    enteredDate = Txt_Data.Text

    If IsDate(enteredDate) Then
        Txt_Data.Text = Format(CDate(enteredDate), "dd/mm/yyyy")
    ElseIf enteredDate <> "" Then
        MsgBox "Invalid date format. Please enter a valid date.", vbExclamation
        Cancel = True ' Prevent focus from leaving the TextBox
    End If
End Sub

Private Sub btn_cancel_Click()

    Unload Me

End Sub

Private Sub Btn_Save_Click()
    
    Dim ws As Worksheet: Set ws = Worksheets("Cadastro")
    Dim dadosOrcto As ListObject: Set dadosOrcto = ws.ListObjects("DadosOrcto")
    
    Call DesbloquearPlanilha

    dadosOrcto.DataBodyRange.Cells(1, 1) = Txt_Cliente.Value
    dadosOrcto.DataBodyRange.Cells(1, 2) = Format(Txt_Data.Text, "mm/dd/yyyy")
    dadosOrcto.DataBodyRange.Cells(1, 3) = Txt_Orcto.Value
    dadosOrcto.DataBodyRange.Cells(1, 4) = Cmb_Tabela.Value

    Unload Me

    Call FormatarCabecalho

    Call BloquearPlanilha

End Sub

