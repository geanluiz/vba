VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private WithEvents tb As MSForms.textBox

Sub Init(tbox As Object)
    Set tb = tbox
End Sub

Private Sub tb_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)

	' Make sure dots are accepted as decimal point
	If InStr(tb.Value, ".") Then
		tb.Value = Replace(tb.Value, ".", ",")
	End If

	' Warns user if he may have not entered a size in centimeters
	If InStr(tb.Value, ",") < 3 And InStr(tb.Value, ",") > 0 Then
		MsgBox "Medida inválida para " & tb.ControlTipText & "." _
			& vbCrLf & "Este campo aceita somente medidas em centímetros.", vbExclamation, "Erro"
	End If
End Sub