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

Private colTB As Collection

Private Sub btn_cancel_Click()

    Unload Me

End Sub

Private Sub medida_padrao()

    Dim tb As Integer
    Dim textBoxes As Variant

    textBoxes = Array(TextLarg, TextProf)


    If btn_medida_padrao.Value Then

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
        Case "Bwc Branco": larg = 80
        Case "Bwc Azul": larg = 115
        Case "Bwc Verde": larg = 70
        Case "Bwc Cinza": larg = 60
    End Select

    TextLarg.Value = larg
    TextProf.Value = 50
    TextAltRod.Value = 7

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

Private Sub btn_ok_Click()

    Dim modelo As String
    Dim largT As Single
    Dim profT As Single
    Dim Valor As Single
    Dim cor As String
    Dim rod As String
    Dim esp As Integer
    Dim ctrl As Object
    Dim cuba As String
    Dim altRod As Single
    Dim qtRod As Integer


    modelo = ComboBox_modelo.Value

    cor = ComboBoxCor.Value

    largT = TextLarg.Value / 100
    profT = TextProf.Value / 100
    altRod = TextAltRod.Value / 100

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.GroupName = "esp" Then
                If ctrl.Value Then
                    esp = 2 + (ctrl.TabIndex * 2)
                End If
            End If

            If ctrl.GroupName = "cuba" Then
                If ctrl.Value Then
                    cuba = ctrl.Name
                End If
            End If
        End If
    Next

    rod = Combo_box_Rod.Value
    
    If rod = "Somente Fundo" Then
        qtRod = 1
    Else
        qtRod = 2
    End If

    
    Valor = Application.Ceiling(vTampo(largT, profT, cor, esp, altRod, qtRod, cuba), 5)

    Call InserirLinha(desc_tampo(cor, largT, profT, cuba, altRod), Valor)


    Unload Me

    Call FormatarTabela
    
    Call FormatarTotais
    
End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim Tbl As ListObject: Set Tbl = ws.ListObjects("coresGranito")
    Dim cores
    Dim modelos
    Dim rod


    modelos = Array("Bwc Branco", "Bwc Azul", "Bwc Verde", "Bwc Cinza")

    ComboBox_modelo.List = modelos
    ComboBox_modelo.ListIndex = 0


    cores = Array("Verde Labrador", "Preto São Gabriel", "Cinza Andorinha", "Branco Itaúnas")

    ComboBoxCor.List = cores
    ComboBoxCor.ListIndex = 0


    rod = Array("Fundo e lateral", "Somente fundo")

    Combo_box_Rod.List = rod
    Combo_box_Rod.ListIndex = 0


    Call medida_padrao


    ' Creates clsTxt classes for each text box to add
    ' An event listener that checks the correct use of floating points
    Dim c As Object
    Set colTB = New Collection
    For Each c In Me.FrameTampo.Controls
        If TypeName(c) = "TextBox" Then
            colTB.Add TbHandler(c)
        End If
    Next c

End Sub

Private Function TbHandler(tb As Object) As ClsTxt
    ' Instantiate objects
    Dim o As New ClsTxt
    o.Init tb
    Set TbHandler = o
End Function
