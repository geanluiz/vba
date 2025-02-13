VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Informe os dados do módulo que deseja inserir"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_cancel_Click()

    Unload Me

End Sub

Private Sub btn_cor_padrao_Click()

    Dim cor As String

    If ComboBox_modelo.Value = "Branco" Then
        cor = "Branca"
    Else
        cor = ComboBox_modelo.Value
    End If
        
    If btn_cor_padrao.Value = True Then
        Text_cor.Value = cor
        Text_cor.ForeColor = &H80000010
        Text_cor.Locked = True
    Else
        Text_cor.ForeColor = &H80000012
        Text_cor.Locked = False
    End If
    
End Sub

Private Sub medida_padrao()

    Dim tb As Integer
    Dim textBoxes As Variant

    textBoxes = Array(TextASup, TextAInf, TextLSup, TextLInf, TextPSup)


    If btn_medida_padrao.Value = True Then

        Call set_medidas

        For tb = LBound(textBoxes) To UBound(textBoxes)
            textBoxes(tb).Locked = True
            textBoxes(tb).ForeColor = &H80000010
        Next tb

        btn_prof50.Visible = True
        btn_prof40.Visible = True
        TextPInf.Visible = False
    Else

        For tb = LBound(textBoxes) To UBound(textBoxes)
            textBoxes(tb).Locked = False
            textBoxes(tb).ForeColor = &H80000012
        Next tb

        btn_prof50.Visible = False
        btn_prof40.Visible = False
        TextPInf.Visible = True

    End If

End Sub

Private Sub set_medidas()

    Dim larg As Single

    Select Case ComboBox_modelo.Value
        Case "Branco": larg = 80
        Case "Azul": larg = 115
        Case "Verde": larg = 70
        Case "Cinza": larg = 60
    End Select

    TextASup.Value = 80
    TextAInf.Value = 70
    TextPSup.Value = 17
    TextLSup.Value = larg
    TextLInf.Value = larg
    TextPInf.Value = 50

End Sub

Private Sub btn_medida_padrao_Click()

    Call medida_padrao

End Sub

Private Sub btn_ok_Click()

    Dim modelo As String
    Dim lSup As Single
    Dim lInf As Single
    Dim aSup As Single
    Dim aInf As Single
    Dim pSup As Single
    Dim pInf As Single
    Dim mold As String
    Dim Valor As Single
    Dim cor As String
    Dim qtPInf As Integer
    Dim i
    
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")

    modelo = ComboBox_modelo.Value

    dict.Add "lSup", Array(CSng(TextLSup / 100), "Largura Superior")
    dict.Add "lInf", Array(CSng(TextLInf / 100), "Largura Inferior")
    dict.Add "aSup", Array(CSng(TextASup / 100), "Altura Superior")
    dict.Add "aInf", Array(CSng(TextAInf / 100), "Altura Inferior")
    dict.Add "pSup", Array(CSng(TextPSup / 100), "Profundidade Superior")

    lSup = dict("lSup")(0)
    lInf = dict("lInf")(0)
    aSup = dict("aSup")(0)
    aInf = dict("aInf")(0)
    pSup = dict("pSup")(0)


    ' Tests if inputs contains wrong separator
    For Each i in dict.keys
        If dict(i)(0) < 0.1 Then
            ' TODO: add return to userform after error message
            MsgBox "Medida inválida para " & dict(i)(1), vbExclamation, "Erro!"
        End If
    Next i


    If btn_medida_padrao.Value = False Then
        pInf = TextPInf.Value / 100
    Else
        If btn_prof50.Value = True Then
            pInf = 0.5
        Else
            pInf = 0.4
        End If
    End If

    If porta_aplq.Value = True Then
        mold = "aplq"
    ElseIf porta_mold.Value = True Then
        mold = "mold"
    Else
        mold = "rpd"
    End If
    

    qtPInf = qtdePortas.Value
    cor = Text_cor.Value
    

    Valor = Application.Ceiling(vBanheiros(modelo, lSup, lInf, aSup, aInf, pSup, pInf, mold, qtPInf), 5)

    Call InserirLinha(desc_text(lSup, lInf, aSup, aInf, pSup, pInf, cor, mold, qtPInf), Valor)

Done:
    Unload Me

    Call FormatarTabela
    
    Call FormatarTotais

    Exit Sub
err:
    MsgBox "Erro! Verifique as informações digitadas e tente novamente...", "Erro!", vbExclamation

End Sub

Private Sub ComboBox_modelo_Change()
    
    Dim cor As String

    If ComboBox_modelo.Value = "Branco" Then
        cor = "Branca"
    Else
        cor = ComboBox_modelo.Value
    End If

    If btn_cor_padrao.Value = True Then
        Text_cor.Value = cor
    Else
        Text_cor.Locked = False
    End If

    Call set_medidas
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim modelos

    modelos = Array("Branco", "Azul", "Verde", "Cinza")
        
    ComboBox_modelo.List = modelos
    ComboBox_modelo.ListIndex = 0

    Text_cor.ForeColor = &H80000010

    Call set_medidas
        
    TextPInf.Visible = False

    Call medida_padrao

    qtdePortas.Value = 2

End Sub


