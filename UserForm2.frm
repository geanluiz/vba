VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Informe os dados do tampo que deseja inserir"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_cancel_Click()

    Unload Me

End Sub

Private Sub medida_padrao()

    Dim tb As Integer
    Dim textBoxes As Variant

    textBoxes = Array(TextLarg, TextProf, TextEsp)


    If btn_medida_padrao.Value = True Then

        Call set_medidas

        ' Change combo boxes font colors to appear as disabled
        For tb = LBound(textBoxes) To UBound(textBoxes)
            textBoxes(tb).Locked = True
            textBoxes(tb).ForeColor = &H80000010
        Next tb

    Else

        ' Change combo boxes font colors to appear as enabled
        For tb = LBound(textBoxes) To UBound(textBoxes)
            textBoxes(tb).Locked = False
            textBoxes(tb).ForeColor = &H80000012
        Next tb

    End If

End Sub

Private Sub set_medidas()

    Dim larg As Single

    Select Case ComboBox_modelo.Value
        Case "Bwc Branco": larg = 0.8
        Case "Bwc Azul": larg = 1.15
        Case "Bwc Verde": larg = 0.7
        Case "Bwc Cinza": larg = 0.6
    End Select

    TextLarg.Value = larg
    TextProf.Value = 0.5
    TextEsp.Value = 2

End Sub

Private Sub btn_medida_padrao_Click()

    Call medida_padrao

End Sub

Private Sub ComboBox_modelo_Change()
    
    Dim cor As String

    If ComboBox_modelo.Value = "Branco" Then
        cor = "Branca"
    Else
        cor = ComboBox_modelo.Value
    End If

    Call set_medidas
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet: Set ws = Worksheets("Cadastro")
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("coresGranito")
    Dim cores
    Dim modelos
    Dim rod

    modelos = Array("Bwc Branco", "Bwc Azul", "Bwc Verde", "Bwc Cinza")

    ComboBox_modelo.List = modelos
    ComboBox_modelo.ListIndex = 0

    cores = Tbl.DataBodyRange
        
    ComboBoxCor.List = cores
    ComboBoxCor.ListIndex = 0

    rod = Array("Somente fundo", "Fundo e lateral")

    Combo_box_Rod.List = rod
    Combo_box_Rod.ListIndex = 0


    Call medida_padrao

    qtdePortas.Value = 2

End Sub
