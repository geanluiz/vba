VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Inserir M�dulo"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
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
        Text_cor.ForeColor = &H80000010&
        Text_cor.Locked = True
    Else
        Text_cor.ForeColor = &H80000012&
        Text_cor.Locked = False
    End If
    
End Sub

Private Sub medida_padrao()

    Dim tb As Integer
    Dim textBoxes As Variant

    textBoxes = Array(TextASup, TextAInf, TextLSup, TextLInf, TextPSup)


    If btn_medida_padrao.Value = True Then

        Call set_medidas

        For tb = LBound(textBoxes) to UBound(textBoxes)
            textBoxes(tb).Locked = True
            textBoxes(tb).ForeColor = &H80000010&
        Next tb        

        btn_prof50.Visible = True
        btn_prof40.Visible = True
        TextPInf.Visible = False
    Else

        For tb = LBound(textBoxes) to UBound(textBoxes)
            textBoxes(tb).Locked = False
            textBoxes(tb).ForeColor = &H80000012&
        Next tb

        btn_prof50.Visible = False
        btn_prof40.Visible = False
        TextPInf.Visible = True

    End If

End Sub

Private Sub set_medidas()

    Dim larg As Single

    Select Case ComboBox_modelo.Value
        Case "Branco": larg = 0.8
        Case "Azul": larg = 1.15
        Case "Verde": larg = 0.7
        Case "Cinza": larg = 0.6
    End Select

    TextASup.Value = 0.8
    TextAInf.Value = 0.7
    TextPSup.Value = 0.17
    TextLSup.Value = larg
    TextLInf.Value = larg
    TextPInf.Value = 0.5

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
    
    modelo = ComboBox_modelo.Value

    lSup = TextLSup.Value
    lInf = TextLInf.Value
    aSup = TextASup.Value
    aInf = TextAInf.Value
    pSup = TextPSup.Value

    If btn_medida_padrao.Value = False Then 
        pInf = TextPInf.Value
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
    
    'qtPortas = InputBox("")
    
    Valor = WorksheetFunction.Ceiling(vBanheiros(modelo, lSup, lInf, aSup, aInf, pSup, pInf, mold), 5)
    cor = Text_cor.Value

    Call InserirLinha(desc_text(lSup, lInf, aSup, aInf, pSup, pInf, cor, mold), Valor)

    Unload Me

    Call FormatarTabela
    
    Call FormatarTotais

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

    Text_cor.ForeColor = &H80000010&

    Call set_medidas
        
    TextPInf.Visible = False

    Call medida_padrao

End Sub

Private Sub ComboBox_modelo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btn_ok_Click
    End If
End Sub


